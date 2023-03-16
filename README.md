# BlenderBIM Spreadsheet

BlenderBIM Spreadsheet was developed and tested with Blender 3.4 on Windows 11 with the lastest stable BlenderBIM release
## Installation

Prequisites

Blender 3.4\
BlenderBIM add-on\
Libre Office Calc\

## Setup 

### Setup of Libre Office Calc
It's highly recommended to use Libre Office Calc when using .ods files.
Download a macro here which creates a table style and autofilter on startup
## Creating your own custom Macro which creates an autofilter and table style

1. Press Record Macro
2. Select cell A1 and Press Ctrl+Shift+Down Arrow then Right Arrow
3. Go to > Format -> Autoformat Styles > Box List Blue > OK
4. Press Home
5. Press Ctrl + Shift + L
6. Stop Recording Macro and Save it
7. Go to Advanced to Uncheck > Java Runtime
8. Make the Macro a hotkey

## About filtering in Libre Office Calc

1. Advanced filters
2. Saving filter
3. Resetting filters


## Functionality

- Support for .xlsx and .ods files
- Users are able to create a Selection Set of Custom Properties which can be saved and shared, similar to Navisworks Search Sets
- Support for all BaseQuantities
- Keeps track of what the user is selecting in the user interface, to prevent multiple spreadsheet files.
- UI of the most commonly used properties such as IsExternal, LoadBearing and FireRating.