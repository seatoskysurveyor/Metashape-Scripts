#Postprocessing Scrip V2.2 Created 6th December 2023
#Agisoft Metashape Professional Version minimum 2.0
#Author : Andreas Taylor, seatoskysurveyor@gmail.com;
#This script provides utilities to process DJI RTK geotagged images
#Ensure project & image folder and subfolder/file permission is set to read and write
#Refer to accompanying "read me" for setup instructions
#Agisoft Preferences must be configured properly e.g. log files, enable python rich console, load image orientations from XML, etc...
#Please use .csv format for marker import

import os
import Metashape
import glob
import csv
import subprocess
from PySide2 import QtGui, QtCore, QtWidgets


# Global Metashape document
#doc = Metashape.app.document

# Define parent globally
app = QtWidgets.QApplication.instance()
parent = app.activeWindow()

# Prompt the user to select the output folder
#output_folder = Metashape.app.getExistingDirectory("Select Output Folder")

# Define the project path globally
#project_path = Metashape.app.document.path

# Define path to photos globally
# path_photos = project_path

# Define the minimum required version
MIN_REQUIRED_VERSION = "2.0.0"

def check_metashape_version():
    # Get the current version of Metashape
    current_version = Metashape.app.version
    # Extract the major, minor, and patch numbers
    major, minor, patch = map(int, current_version.split('.')[:3])
    required_major, required_minor, required_patch = map(int, MIN_REQUIRED_VERSION.split('.'))
    
    # Check if the current version is equal to or higher than the required version
    if (major, minor, patch) < (required_major, required_minor, required_patch):
        print(f"Your Metashape version {current_version} is not supported by this script.")
        print(f"Please upgrade to Metashape version {MIN_REQUIRED_VERSION} or later.")
        return False
    return True
    
# Function to add altitude to camera reference
def add_altitude():
    """
    Adds user-defined altitude for camera instances in the Reference pane
    """
    doc = Metashape.app.document
    if not len(doc.chunks):
        raise Exception("No chunks!")

    alt = Metashape.app.getFloat("Please specify the height to be added:", 100)
    print("Script started...")
    chunk = doc.chunk

    for camera in chunk.cameras:
        if camera.reference.location:
            coord = camera.reference.location
            camera.reference.location = Metashape.Vector([coord.x, coord.y, coord.z + alt])

    print("Script finished!")


# Function to read DJI relative altitude
def read_DJI_relative_altitude():
    """
    Reads DJI/RelativeAltitude information from the image meta-data and writes it to the Reference pane
    """
    doc = Metashape.app.document
    if not len(doc.chunks):
        raise Exception("No chunks!")

    print("Script started...")
    chunk = doc.chunk

    for camera in chunk.cameras:
        if not camera.type == Metashape.Camera.Type.Regular: #skip camera track, if any
            continue
        if not camera.reference.location:
            continue
        if "DJI/RelativeAltitude" in camera.photo.meta.keys():
            z = float(camera.photo.meta["DJI/RelativeAltitude"])
            camera.reference.location = Metashape.Vector([camera.reference.location.x, camera.reference.location.y, z])

    print("Script finished!")    

def open_project_folder(project_folder):
    if os.name == 'posix':  # macOS and Linux
        subprocess.Popen(['open', project_folder])
    elif os.name == 'nt':  # Windows
        os.startfile(project_folder)
        
def get_images_folder():
    app = QtWidgets.QApplication.instance()
    parent = app.activeWindow()
    
    while True:
        # Prompting user to select folder
        images_folder = QtWidgets.QFileDialog.getExistingDirectory(parent, "Select the Images Folder")
        
        if not images_folder:
            print("Image folder selection cancelled.")
            return None

        # Check if the folder is likely on an external drive
        if ":\\" in images_folder or "/Volumes/" in images_folder:
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setText("The selected folder appears to be on an external drive (like an SD card).")
            msg_box.setInformativeText("Processing directly from an external drive can be slower and risk data loss if the drive is disconnected. Do you want to continue?")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            msg_box.setDefaultButton(QtWidgets.QMessageBox.No)
            reply = msg_box.exec_()
            
            if reply == QtWidgets.QMessageBox.Yes:
                break  # User chose to proceed
            else:
                print("Please select a different folder.")
                # Loop back to folder selection
        else:
            break  # Folder is likely not on an external drive

    return images_folder
    
# Define the function to import flight boundary shape for clipping export data     
def import_KML():

    # Open file dialog to select KML file
    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(parent, "Select KML File", "", "KML Files (*.kml)")
    if not file_path:
        print("File selection cancelled.")
        return

    # Get the active chunk
    doc = Metashape.app.document
    if doc.chunk is None:
        print("No active chunk found.")
        return

    try:
        # Import KML file
        doc.chunk.importShapes(path=file_path)
        print(f"KML file '{os.path.basename(file_path)}' imported successfully.")
    except Exception as e:
        print(f"Error importing KML file: {e}")

# Define the function to import images
def import_images_from_folders(images_folder):
    print("Importing images from:", images_folder)
    try:
        chunk = create_or_get_chunk()
        if chunk:
            extensions = ['jpg', 'tif', 'png']
            images = []
            for ext in extensions:
                # Constructing the file pattern for glob
                images_pattern = os.path.join(images_folder, '**', f'*.{ext}')
                found_images = glob.glob(images_pattern, recursive=True)
                images.extend(found_images)

            if images:
                chunk.addPhotos(images)
                print("Images imported successfully.")
            else:
                print("No images found to import.")
    except Exception as e:
        print(f"Error importing images: {e}")
        
# Define create or get existing chunk        
def create_or_get_chunk():
    try:
        if not doc.chunks:
            chunk = doc.addChunk()
            chunk.label = "Chunk_1"
        else:
            chunk = doc.chunk

        return chunk
    except Exception as e:
        print(f"Error creating/getting chunk: {e}")
        return None
 
# Define the function to convert coordinates of cameras and markers
def convert_reference(chunk, target_crs):
    original_crs = chunk.crs
    for camera in chunk.cameras:
        if camera.reference.location:
            camera.reference.location = Metashape.CoordinateSystem.transform(camera.reference.location, original_crs, target_crs)
    for marker in chunk.markers:
        if marker.reference.location:
            marker.reference.location = Metashape.CoordinateSystem.transform(marker.reference.location, original_crs, target_crs)
    chunk.crs = target_crs
    chunk.updateTransform()
    
# Define the function to align the images
def align_images(chunk):
    try:
        # Perform image alignment
        chunk.alignPhotos()

        print("Image alignment completed.")
    except Exception as e:
        print(f"Error aligning images: {e}")       

# Define the function to import ground control points/check points   
def import_reference_markers(chunk):
    app = QtWidgets.QApplication.instance()
    parent = app.activeWindow()

    # Open file dialog to select CSV or TXT file
    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(parent, "Select CSV/TXT File", "", "CSV/TXT Files (*.csv *.txt)")
    if not file_path:
        print("File selection cancelled.")
        return

    try:
        # Read the file and add markers
        with open(file_path, 'r') as file:
            for line in file:
                # Skip empty lines
                if not line.strip():
                    continue

                # Split the line by comma or whitespace
                row = line.strip().split(',')
                if len(row) < 4:
                    row = line.strip().split()

                if len(row) != 4:
                    print("Skipping row: Incorrect number of values. Expected 4 columns (Name, Easting, Northing, Altitude).")
                    continue

                # Parse the marker data
                name, easting, northing, altitude = row
                x, y, z = float(easting), float(northing), float(altitude)

                # Create a marker in the chunk
                marker = chunk.addMarker()
                marker.label = name
                marker.reference.location = Metashape.Vector([x, y, z])
                marker.reference.enabled = True

        print(f"File '{os.path.basename(file_path)}' imported successfully.")
    except Exception as e:
        print(f"Error importing file: {e}")

        
#downscale set to 8 for "low quality", original value=4:
def generate_depth_maps(chunk):
    try:
        chunk.buildDepthMaps(downscale=8, filter_mode=Metashape.MildFiltering) 
    except Exception as e:
        print(f"Error generating depth maps: {e}")

#Generate point cloud with point confidences:
def generate_point_cloud(chunk):
    try:
        if chunk.transform.scale and chunk.transform.rotation and chunk.transform.translation:
            chunk.buildPointCloud(point_confidence=True)
        else:
            print("Error: Transformation not applied. Please apply transformation first.")
    except Exception as e:
        print(f"Error generating point cloud: {e}")

def generate_dem(chunk):
    try:
        if chunk.point_cloud:
            chunk.buildDem(source_data=Metashape.PointCloudData)
        else:
            print("Error: Point cloud not available. Please generate point cloud first.")
    except Exception as e:
        print(f"Error generating DEM: {e}")
        
def generate_orthomosaic(chunk):
    try:
        if chunk.elevation:
            chunk.buildOrthomosaic(surface_data=Metashape.ElevationData, resolution=0.1)
        else:
            print("Error: Elevation data not available. Please generate DEM first.")
    except Exception as e:
        print(f"Error generating orthomosaic: {e}")

def export_las(chunk):
    try:
        if chunk.point_cloud:
            project_name = Metashape.app.document.path.split("/")[-1].split(".")[0]
            output_folder = Metashape.app.document.path.rsplit("/", 1)[0]
            # Change the extension to .laz for compressed format
            laz_output_path = os.path.join(output_folder, f'{project_name}_point_cloud.laz')
            # Assuming Metashape supports .laz format in the exportPointCloud method
            chunk.exportPointCloud(laz_output_path, source_data=Metashape.PointCloudData, clip_to_boundary=True)
        else:
            print("Error: Point cloud not available. Please generate point cloud first.")
    except Exception as e:
        print(f"Error exporting LAZ file: {e}")

def export_orthomosaic(chunk, format='png', resolution=0.1):
    try:
        if chunk.point_cloud:
            project_name = Metashape.app.document.path.split("/")[-1].split(".")[0]
            output_folder = Metashape.app.document.path.rsplit("/", 1)[0]
            ortho_output_path = os.path.join(output_folder, f'{project_name}_ortho_{int(resolution*100)}cm.{format}')
            chunk.exportRaster(ortho_output_path, format=Metashape.RasterFormatTiles, 
                               image_format=Metashape.ImageFormatPNG if format == 'png' else Metashape.ImageFormatTIFF,
                               raster_transform=Metashape.RasterTransformNone, save_world=True, clip_to_boundary=True, resolution=resolution)
        else:
            print("Error: Orthomosaic not available. Please generate orthomosaic first.")
    except Exception as e:
        print(f"Error exporting orthomosaic: {e}")

def export_report(chunk):
    try:
        project_name = Metashape.app.document.path.split("/")[-1].split(".")[0]
        output_folder = Metashape.app.document.path.rsplit("/", 1)[0]
        report_output_path = os.path.join(output_folder, f'{project_name}_report.pdf')
        chunk.exportReport(report_output_path)
    except Exception as e:
        print(f"Error exporting report: {e}")
        
        
######### ***************Define Workflow to process DJI RTK geotagged images WITHOUT Ground Control Points (GCP's)******************
def workflow_DJI():
    global parent

    # Get the Metashape document
    doc = Metashape.app.document

    # Check if the document has any chunks
    if not doc.chunks:
        print("No active chunks in the Metashape document.")
        return       

    # Now it's safe to get the project path
    project_path = doc.path
   
    if not check_metashape_version():
        return  # Stop execution if the version check fails 
    
    # Get the images folder from the user
    images_folder = get_images_folder()
    if not images_folder:
        print("Image folder selection cancelled.")
        return

    print("Beginning start-up sequence")

    # Call the function to import images from the selected folder
    import_images_from_folders(images_folder)

    print("Checking save filename")
    project_path = Metashape.app.getSaveFileName("Specify project name and location for saving:")
    if not project_path:
        print("Booting Down")
        return 0

    if project_path[-4:].lower() != ".psx":
        project_path += ".psx"

    crs = Metashape.app.getCoordinateSystem("Select Coordinate System")

    # Use the existing chunk
    doc.save(path=project_path)
    chunk = doc.chunk  # Assuming there is always an existing chunk when the script is run

    print("Adding photos")
    try:
        image_list = [
            os.path.join(images_folder, path)
            for path in os.listdir(images_folder)
            if path.lower().endswith(("jpg", "jpeg", "tif", "png", "JPG", "JPEG", "TIF", ""))
        ]

        if not image_list:
            print("No images found to add.")
            return

        chunk.addPhotos(image_list, load_reference=True, load_xmp_calibration=True, load_xmp_orientation=True, load_xmp_accuracy=True, load_xmp_antenna=True)

        # Process cameras
        for camera in chunk.cameras:
            if camera.reference.location:
                coord = camera.reference.location
                camera.reference.location = Metashape.CoordinateSystem.transform(coord, chunk.crs, crs)
        chunk.crs = crs     
                
    except Exception as e:
        print(f"Error in adding photos or processing cameras: {e}")      
    
    # Convert reference crs
    if not doc.chunks:
        Metashape.app.messageBox("No chunk is available.")
    else:
        chunk = doc.chunk
        target_crs = Metashape.app.getCoordinateSystem("Select Target Coordinate System", chunk.crs)
        if target_crs:
            convert_reference(chunk, target_crs)
            print("Coordinate system conversion completed.")
        else:
            print("Coordinate system selection cancelled.")
           
    
    print ("Align Photos")
    #chunk.crs = crs
    chunk.matchPhotos(downscale=1, generic_preselection=True, reference_preselection=True,filter_mask=False, keypoint_limit=40000, tiepoint_limit=4000)
    chunk.alignCameras(adaptive_fitting=False)
        
    #document save
    doc.save()

    #Optimize Cameras
    chunk.optimizeCameras(fit_f=True, fit_cx=True, fit_cy=True, fit_b1=True, fit_b2=True, fit_k1=True,fit_k2=True, fit_k3=True, fit_k4=False, fit_p1=True, fit_p2=True, fit_p3=False,fit_p4=False, fit_corrections=True, tiepoint_covariance=True)

    #Set Coordinate System
    chunk.updateTransform()
    doc.save()

    # Generate depth maps
    generate_depth_maps(chunk)

    # Generate point cloud
    generate_point_cloud(chunk)

    #save project
    doc.save()

    # Generate DEM
    generate_dem(chunk)

    # Generate orthomosaic
    generate_orthomosaic(chunk)
    #document save
    doc.save()
        
    # Construct the base name for output files using project_path
    base_output_name = os.path.splitext(os.path.basename(project_path))[0]   
        
    # Export results

    # Export functions using project_folder for saving the files
    # Export results
    export_las(chunk)  # Pass only the chunk as the function constructs the file path internally
    export_orthomosaic(chunk, format='png', resolution=0.2)
    export_orthomosaic(chunk, format='tif', resolution=0.04)
    export_orthomosaic(chunk, format='png', resolution=0.3)
    export_report(chunk)

    # Determine the project folder
    if doc.path:  # Check if the project is saved
        project_folder = os.path.dirname(doc.path)
    else:
        print("Project not saved. Please save the project first.")
        return 

    print(f"Processing finished, results saved to {project_folder}.")
    open_project_folder(project_folder)
    
    
       
##********Define Workflow to process RTK/PPK geotagged images WITH Ground Control Points (GCP)Step 1 stops for user to refine markers in GUI********
def workflow_DJI_step1():
    global parent

    # Get the Metashape document
    doc = Metashape.app.document

    # Check if the document has any chunks
    if not doc.chunks:
        print("No active chunks in the Metashape document.")
        return

    # Now it's safe to get the project path
    project_path = doc.path
    
    if not check_metashape_version():
        return  # Stop execution if the version check fails 
    
    # Get the images folder from the user
    images_folder = get_images_folder()
    if not images_folder:
        print("Image folder selection cancelled.")
        return

    print("Beginning start-up sequence")

    # Call the function to import images from the selected folder
    import_images_from_folders(images_folder)

    print("Checking save filename")
    project_path = Metashape.app.getSaveFileName("Specify project name and location for saving:")
    if not project_path:
        print("Booting Down")
        return 0

    if project_path[-4:].lower() != ".psx":
        project_path += ".psx"

    crs = Metashape.app.getCoordinateSystem("Select Coordinate System")

    # Use the existing chunk
    doc.save(path=project_path)
    chunk = doc.chunk  # Assuming there is always an existing chunk when the script is run

    print("Adding photos")
    try:
        image_list = [
            os.path.join(images_folder, path)
            for path in os.listdir(images_folder)
            if path.lower().endswith(("jpg", "jpeg", "tif", "png", "JPG", "JPEG", "TIF", ""))
        ]

        if not image_list:
            print("No images found to add.")
            return

        chunk.addPhotos(image_list, load_reference=True, load_xmp_calibration=True, load_xmp_orientation=True, load_xmp_accuracy=True, load_xmp_antenna=True)

        # Process cameras
        for camera in chunk.cameras:
            if camera.reference.location:
                coord = camera.reference.location
                camera.reference.location = Metashape.CoordinateSystem.transform(coord, chunk.crs, crs)
        chunk.crs = crs     
                
    except Exception as e:
        print(f"Error in adding photos or processing cameras: {e}")      
    
    # Convert reference crs
    if not doc.chunks:
        Metashape.app.messageBox("No chunk is available.")
    else:
        chunk = doc.chunk
        target_crs = Metashape.app.getCoordinateSystem("Select Target Coordinate System", chunk.crs)
        if target_crs:
            convert_reference(chunk, target_crs)
            print("Coordinate system conversion completed.")
        else:
            print("Coordinate system selection cancelled.")
             
   
    # Call the function to import markers
    import_reference_markers(chunk)
    
    print ("Align Photos")
    #chunk.crs = crs
    chunk.matchPhotos(downscale=1, generic_preselection=True, reference_preselection=True,filter_mask=False, keypoint_limit=40000, tiepoint_limit=4000)
    chunk.alignCameras(adaptive_fitting=False)
        
    #document save
    doc.save()

    #Optimize Cameras
    #chunk.optimizeCameras(fit_f=True, fit_cx=True, fit_cy=True, fit_b1=True, fit_b2=True, fit_k1=True,fit_k2=True, fit_k3=True, fit_k4=False, fit_p1=True, fit_p2=True, fit_p3=False,fit_p4=False, fit_corrections=True, tiepoint_covariance=True)

    #Set Coordinate System
    chunk.updateTransform()
    doc.save()

    #chunk.importReference(path=path_markers, format=Metashape.ReferenceFormatCSV, columns= 'nxyz', delimiter=',',ignore_labels=False,create_markers=True)
    msg = "Place flags in the center of the marks."
    reply = QtWidgets.QMessageBox.question(parent, 'Message', msg, QtWidgets.QMessageBox.Ok)
    print("Step 1 workflow finished continue with step 2")
    return 1



#### ************ Define Workflow to process RTK geotagged images WITH Ground Control Points (GCP)Step 2 - the "batch process"	***************
def workflow_DJI_step2():
    global parent
    
    # Get the Metashape document
    doc = Metashape.app.document
    
     # Determine the project folder
    if doc.path:  # Check if the project is saved
        project_folder = os.path.dirname(doc.path)
    else:
        print("Project not saved. Please save the project first.")
        return
        
    # Check if there are any chunks in the document
    if len(doc.chunks) > 0:
        chunk = doc.chunks[0]  # Access the first chunk
        print("Workflow Step 2:\nStarting processing...")
        # Proceed with operations on the chunk
    else:
        print("No chunks found in the document")
        # Handle the situation where no chunks are available

    # Now it's safe to get the project path
    project_path = doc.path
    
    # Optimize Cameras
    chunk.optimizeCameras(fit_f=True, fit_cx=True, fit_cy=True, fit_b1=True, fit_b2=True, fit_k1=True,fit_k2=True, fit_k3=True, fit_k4=False, fit_p1=True, fit_p2=True, fit_p3=False,fit_p4=False)

    # Set Coordinate System
    chunk.updateTransform()
    chunk.resetRegion()
    doc.save()

    # Generate depth maps
    generate_depth_maps(chunk)

    # Generate point cloud
    generate_point_cloud(chunk)

    # Save project
    doc.save()

    # Generate DEM
    generate_dem(chunk)

    # Generate orthomosaic
    generate_orthomosaic(chunk)
    # Document save
    doc.save()
        
    # Construct the base name for output files using project_path
    base_output_name = os.path.splitext(os.path.basename(project_path))[0]   
        
    # Export results

    # Export functions using project_folder for saving the files
    # Export results
    export_las(chunk)  # Pass only the chunk as the function constructs the file path internally
    export_orthomosaic(chunk, format='png', resolution=0.2)
    export_orthomosaic(chunk, format='tif', resolution=0.04)
    export_orthomosaic(chunk, format='png', resolution=0.3)
    export_report(chunk)

    print(f"Processing finished, results saved to {project_folder}.")
    open_project_folder(project_folder)
    
    print("Step 2 workflow finished.")
    return 1      

# Add the workflows to Metashape menu:

label_import_kml = "DJI Utilities/Import KML Clipping Boundary"
Metashape.app.addMenuItem(label_import_kml, import_KML)
print("To execute the 'Import KML' tool, go to: {}".format(label_import_kml))

label = "DJI Utilities/Process RTK without GCP"
Metashape.app.addMenuItem(label, workflow_DJI)
print("To execute this script press {}".format(label))

label = "DJI Utilities/ Process RTK with GCP - Step 1"
Metashape.app.addMenuItem(label, workflow_DJI_step1)
print("To execute this script press {}".format(label))

label = "DJI Utilities/ Process RTK with GCP - Step 2"
Metashape.app.addMenuItem(label, workflow_DJI_step2)
print("To execute this script press {}".format(label))

label = "DJI Utilities/Add Altitude to Camera Reference"
Metashape.app.addMenuItem(label, add_altitude)
print("To execute the 'Add Altitude' script, go to: {}".format(label))

label = "DJI Utilities/Read Relative Altitude from DJI Metadata"
Metashape.app.addMenuItem(label, read_DJI_relative_altitude)
print("To execute the 'Read DJI Altitude' script, go to: {}".format(label))

