import bpy
from bpy.types import Panel

from . import  prop, operators

#initial hardcoded global variables
#my_ifcproduct = 'IfcProduct'

class BlenderBIMSpreadSheetPanel(Panel):
    bl_idname = "OBJECT_PT_BlenderBIMSpreadSheet_panel"
    bl_label = "BlenderBIM Spreadsheet"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "BlenderBIM | Spreadsheet"

    def draw(self, context):

        ifc_properties = context.scene.ifc_properties
        layout = self.layout
        
        layout.label(text="General")
        box = layout.box()
        row = box.row()
        row.prop(ifc_properties, "my_ifcproduct")

        row = box.row()
        row.prop(ifc_properties, "my_ifcproductname")

        row = box.row()
        row.prop(ifc_properties, "my_ifcproducttypename")

        row = box.row()
        row.prop(ifc_properties, "my_ifcbuildingstorey")

        row = box.row()
        row.prop(ifc_properties, "my_ifcclassification")

        row = box.row()
        row.prop(ifc_properties, "my_ifcmaterial" )


        layout.label(text="Common Properties")
        box = layout.box()

        row = box.row()
        row.prop(ifc_properties, "my_isexternal")

        row = box.row()
        row.prop(ifc_properties, "my_loadbearing")

        row = box.row()
        row.prop(ifc_properties, "my_firerating")

        #row = box.row()
        #row.prop(ifc_properties, "my_ifcclassification_dd")

        
        layout.label(text="Custom Properties")
        box = layout.box()

        custom_collection = context.scene.custom_collection

        row = layout.row(align=True)
        row.operator("custom.collection_actions", text="Add", icon="ADD").action = "add"
        row.operator("custom.collection_actions", text="Remove Last", icon="REMOVE").action = "remove"

        for item in custom_collection.items:
            box.prop(item, "name")







        layout.label(text="Write to Spreadsheet")
        box_spreadsheet = layout.box()

        row = box_spreadsheet.row()
        row.prop(ifc_properties, "my_spreadsheetfile")

        row = box_spreadsheet.row()
        row.prop(ifc_properties, "ods_or_xlsx")

        
        
        box_spreadsheet.operator("export.tospreadsheet")





        layout.label(text="Filter IFC elements")

      
        self.layout.operator("object.filter_ifc_elements", text="Filter IFC elements", icon="FILTER")
        self.layout.operator("object.unhide_all", text="Unhide IFC elements", icon="LIGHT")

def register():
    bpy.utils.register_class(BlenderBIMSpreadSheetPanel)

def unregister():
    bpy.utils.unregister_class(BlenderBIMSpreadSheetPanel)

"""    
class BlenderBIMSpreadSheetPanel(Panel):
    bl_label = "BlenderBIM Spreadsheet"
    bl_idname = "OBJECT_PT_blenderbimxlsxpanel"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "Tools"

    def draw(self, context):
        row = self.layout.row()
        row.prop(context.scene, 'property_ifcproduct')


def register():
    bpy.utils.register_class(BlenderBIMSpreadSheetPanel)

def unregister():
    bpy.utils.unregister_class(BlenderBIMSpreadSheetPanel)


class BlenderBIMSpreadSheetPanel(Panel):
    bl_label = "BlenderBIM spreadsheet"
    bl_idname = "OBJECT_PT_blenderbimxlsxpanel"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "Tools"

    def draw(self, context):
        row = self.layout.row()
        row.prop(context.scene, 'my_property')
        row.prop(context.scene, 'another_property')
        row.prop(context.scene, 'yet_another_property')
        row.prop(context.scene, 'and_yet_another_property')

def register():
    bpy.utils.register_class(BlenderBIMSpreadSheetPanel)

def unregister():
    bpy.utils.unregister_class(BlenderBIMSpreadSheetPanel)

 
import os
import sys
import time
import site
import collections
import subprocess

import bpy
from bpy.props import StringProperty, BoolProperty, IntProperty, EnumProperty
from bpy_extras.io_utils import ImportHelper 
from bpy.types import (Operator, PropertyGroup)

from . import  prop, operator
    
class BlenderBIMSpreadSheetPanel(bpy.types.Panel):
    #Creates a Panel in the Object properties window  

    #bl_label = "BlenderBIM spreadsheet"
    #bl_idname = "OBJECT_PT_blenderbimxlsxpanel"  # this is not strictly necessary
    #bl_space_type = "VIEW_3D"
    #bl_region_type = "UI"
    #bl_category = "BlenderBIM"
    
    bl_label = "Spreadsheet Writer"
    bl_options = {"DEFAULT_CLOSED"}
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "BlenderBIM"
 
    def draw(self, context):
        
        layout = self.layout
        blenderbim_spreadsheet_properties = context.scene.blenderbim_spreadsheet_properties
        
        layout.label(text="General")
        box = layout.box()
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_ifcproduct")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_ifcbuildingstorey")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_ifcproduct_name")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_type")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_ifcclassification")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_ifcmaterial")
        
        layout.label(text="Common Properties")
        box = layout.box()
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_isexternal")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_loadbearing")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_firerating")
        
        layout.label(text="BaseQuantities")
        box = layout.box()
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_length")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_width")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_height")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_area")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_netarea")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_netsidearea")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_volume")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_netvolume")
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_perimeter")
        
        layout.label(text="Custom Properties")
        my_collection = context.scene.my_collection
       
        row = layout.row(align=True)
        row.operator("my.collection_actions", text="Add", icon="ADD").action = "add"
        row.operator("my.collection_actions", text="Remove Last", icon="REMOVE").action = "remove"

        for item in my_collection.items:
            layout.prop(item, "name")
        
        layout.label(text="Write to spreadsheet")
        self.layout.operator(operator.WriteToXLSX.bl_idname, text="Write IFC data to .xlsx", icon="FILE")
        self.layout.operator(operator.WriteToODS.bl_idname, text="Write IFC data to .ods", icon="FILE")
        
        layout.label(text="Filter IFC elements")
  
        box = layout.box()
        row = box.row()
        row.prop(blenderbim_spreadsheet_properties, "my_file_path")
        self.layout.operator(operator.FilterIFCElements.bl_idname, text="Filter IFC elements", icon="FILTER")
        self.layout.operator(operator.UnhideIFCElements.bl_idname, text="Unhide IFC elements", icon="LIGHT")
        
"""