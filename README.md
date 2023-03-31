# BlenderBIM Spreadsheet

BlenderBIM Spreadsheet was developed and tested with Blender 3.5 on Windows 11 with the lastest stable BlenderBIM release.
With the intention that it will be rolled into BlenderBIM.

Prequisites

[Blender 3.5](https://www.blender.org/download/)\
[BlenderBIM add-on latest stable release](https://blenderbim.org/download.html)\
[Libre Office Calc](https://www.libreoffice.org/download/download-libreoffice/)

## Functionality and new features

- Support for .xlsx and .ods files
- You can now export all BaseQuantities such as Area, NetArea, NetSideArea etc.
- You can now store  user interface settings, useful for many custom properties which can be shared through json files.
- Frienldy user interface to export one of the most commonly used properties such as IsExternal, LoadBearing, FireRating and AcousticRating.
- Support for multiple classification references
- Support for multiple material definitions
- Support for IFC4
## Installation



### Setup of Libre Office Calc
It's highly recommended to use Libre Office Calc when using .ods files.

## Quickstart on creating your own custom Macro which creates an autofilter and table style

1. Press Record Macro
2. Select cell A1 and Press Ctrl+Shift+Down Arrow then Right Arrow
3. Go to > Format -> Autoformat Styles > Box List Blue > OK
4. Press Home
5. Press Ctrl + Shift + L
6. Stop Recording Macro and Save it
7. Go to Advanced to Uncheck > Java Runtime
8. Make the Macro a hotkey
