# BlenderBIM Spreadsheet
# Contributor(s): Coen Claus 

# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTIBILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
# General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.
import os
import bpy
import site

from . import operators, ui, prop

bl_info = {
        "name": "BlenderBIM Spreadsheet",
        "description": "Filter IFC elements from a spreadsheet",
        "author": "Coen Claus",
        "version": (1, 0),
        "blender": (3, 4, 1),
        "location": "Tools",
        "warning": "depends on pandas, xlsxwriter and ods",
        "wiki_url": "https://github.com/C-Claus",
        "tracker_url": "https://community.osarch.org/",
        "support": "COMMUNITY",
        "category": "Import-Export"
        }


site.addsitedir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "libs", "site", "packages"))

""" 
classes = (
    ui.BlenderBIMSpreadSheetPanel,
    #prop.BlenderBIMSpreadSheetProperties
)

def register():

    bpy.types.Scene.blenderbim_spreadsheet_properties = bpy.props.PointerProperty(type=prop.BlenderBIMSpreadSheetProperties)    
   

def unregister():
    
    bpy.types.Scene.blenderbim_spreadsheet_properties

""" 
def register():
    from . import prop
    from . import ui
    from . import operators
    prop.register()
    operators.register()
    ui.register()

def unregister():
    from . import prop
    from . import ui
    from . import operators
    prop.unregister()
    operators.register()
    ui.unregister()

if __name__ == '__main__':
    register()


