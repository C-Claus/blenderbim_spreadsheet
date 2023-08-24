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

from . import operator, ui, prop

bl_info = {
        "name": "BlenderBIM Spreadsheet",
        "description": "Filter IFC elements from a spreadsheet",
        "author": "Coen Claus",
        "version": (2, 0),
        "blender": (3, 6, 2),
        "location": "BlenderBIM | Spreadsheet tab",
        "wiki_url": "https://github.com/C-Claus",
        "tracker_url": "https://community.osarch.org/",
        "support": "COMMUNITY",
        "category": "Import-Export"
        }


site.addsitedir(os.path.join(os.path.dirname(os.path.realpath(__file__)), "libs", "site", "packages"))

def register():
    from . import prop
    from . import ui
    from . import operator
    prop.register()
    operator.register()
    ui.register()

def unregister():
    from . import prop
    from . import ui
    from . import operator
    prop.unregister()
    operator.register()
    ui.unregister()

if __name__ == '__main__':
    register()


