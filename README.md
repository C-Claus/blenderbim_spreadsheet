# BlenderBIM Spreadsheet

The [BlenderBIM Spreadsheet add-on](https://github.com/C-Claus/blenderbim_spreadsheet/releases/download/v1.0/blenderbim_spreadsheet.zip) was developed and tested with Blender 3.5 with the BlenderBIMv0.0.230304 release and LibreOffice Calc 7.5 on Windows 11.
The intention of this spreadsheet add-on is that it will be rolled into BlenderBIM one day.

Prequisites

[Blender 3.5](https://www.blender.org/download/)\
[BlenderBIM add-on latest stable release](https://blenderbim.org/download.html)\
[Libre Office Calc](https://www.libreoffice.org/download/download-libreoffice/)

## Functionality and new features

- Support for .xlsx and .ods files
- You can now export all BaseQuantities such as *Area, NetArea, NetSideArea* etc.
- You can now store  user interface settings, useful for many custom properties which can be shared through json files.
- Frienldy user interface to export one of the most commonly used properties such as *IsExternal, LoadBearing, FireRating* and *AcousticRating*.
- Support for multiple classification references including ```IFC2X3``` and ```IFC4```
- Support for multiple material definitions including *IfcMaterial, IfcMaterialList, IfcMaterialConstituentSet, IfcMaterialLayerSetUsage* and *IfcMaterialProfileSet*.
- Support for IFC2x3 and IFC4

## Installation

BlenderBIM Spreadsheet is found under the BlenderBIM Spreadsheet tab

<img src="https://user-images.githubusercontent.com/14906760/229779615-69e27c38-6eee-4c05-8eea-5144d3deb2f0.png"  width="60%" height="30%">

### Usage

The BlenderBIM spreadsheet add-on expects an IFC file to be loaded and saved with BlenderBIM before an export to a spreadsheet is possbile.
When a table filtering is made in the spreadsheet file, it should be saved first before a 3D view filtering in Blender is possible.
In this image an example is shown with a filtering of the IfcWindow, which shows the multiple materials assigned.

<img src="https://user-images.githubusercontent.com/14906760/229783041-59dfa666-82ae-4b44-ad37-0477da0c4638.png"  width="60%" height="30%">

It's possible to store all your custom defined properties in a .json file. It will save your checkbox user interface settings as well.
To export the properties you want define the PropertySet name and Property name divided by a . as shown in the image

<img src="https://user-images.githubusercontent.com/14906760/229784362-04d4a822-6c6b-4fd6-b3e3-73699b0f70cd.png"  width="60%" height="30%">

When you click *Save Selection Set* a .json file will be written to the same location where your ifc is.Before creating a new spreadsheet export it's important to close the already running instance of your spreadsheet software.In this example the *LoadBearing True* property is filtered in the spreadsheet software en visualized in the Blender 3D View.

<img src="https://user-images.githubusercontent.com/14906760/229786541-a6e92852-46cb-43b8-964b-87661fd26e5f.png"  width="60%" height="30%">



### Setup of Libre Office Calc
It's highly recommended to use Libre Office Calc when using .ods files. Experience learns MS excel does not always produce valid xml files which are used to parse the *GlobalId* to filter in a 3D View in Blender.

## Quickstart on creating your own custom Macro which creates an autofilter and table style

1. Press Record Macro
2. Select cell A1 and Press Ctrl+Shift+Down Arrow then Right Arrow
3. Go to > Format -> Autoformat Styles > Box List Blue > OK
4. Press Home
5. Press Ctrl + Shift + L
6. Stop Recording Macro and Save it
7. Go to Advanced to Uncheck > Java Runtime
8. Make the Macro a hotkey
