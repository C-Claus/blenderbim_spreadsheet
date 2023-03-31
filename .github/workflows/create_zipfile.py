import os
import zipfile
from pathlib import Path

addon_zip_list  = []
zip_name        = 'blenderbim_spreadsheet.zip'
zip_object      = zipfile.ZipFile(zip_name, "w")

file_path = os.listdir(Path.cwd())

for addon_file in file_path:
    if addon_file == 'libs':
        addon_zip_list.append(addon_file)
        #zip_object.write((file_path, os.path.relpath(file_path, directory)))
    if addon_file.endswith('.py') and addon_file != 'create_zipfile.py':
        addon_zip_list.append(addon_file)

        #with zipfile.ZipFile('blenderbim_spreadsheet.zip', 'w') as zip_object:
        #zip_object.write(file_path, os.path.relpath(file_path, directory))


zip_object.close()

#with zipfile.ZipFile('blenderbim_spreadsheet.zip', 'w') as zip_object:
   # Traverse all files in directory



"""

absolute_path = os.path.dirname(__file__)
relative_path = "lib"
full_path = os.path.join(absolute_path, relative_path)

with os.scandir(path) as it:
    for entry in it:
        if entry.name.endswith(".py") and entry.is_file():
            print(entry.name, entry.path)   


for folder_name, sub_folders, file_names in os.walk('blenderbim_spreadsheet'):
    #print (folder_name)
    for filename in file_names:
        print ('test', filename)
         # Create filepath of files in directory
   #      file_path = os.path.join(folder_name, filename)
         # Add files to zip file
   #      zip_object.write(file_path, os.path.basename(file_path))

#if os.path.exists('blenderbim_spreadsheet.zip'):
#   print("ZIP file created")
#else:
#   print("ZIP file not created")

"""