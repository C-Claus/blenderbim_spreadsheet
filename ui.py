import bpy
from bpy.types import Panel
from . import  prop, operators

class GENERAL_panel:
    bl_label = "BlenderBIM Spreadsheet"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "BlenderBIM | Spreadsheet"
    bl_options = {"DEFAULT_CLOSED"}

class GENERAL_PT_PANEL(GENERAL_panel, Panel):
    bl_idname = "EXAMPLE_PT_panel_1"
    bl_label = "BlenderBIM | Spreadsheet"

    def draw(self, context):
        layout = self.layout
class GENERAL_IFC_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "General"

    def draw(self, context):

        ifc_properties = context.scene.ifc_properties
        layout = self.layout

        box = layout.box()
        row = box.row()
        box.prop(ifc_properties, "my_ifcproduct")
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
class COMMON_PROPERTIES_IFC_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "Common Properties"

    def draw(self, context):

        ifc_properties = context.scene.ifc_properties

        layout = self.layout
        box = layout.box()
        #row = box.row()

        row = box.row()
        row.prop(ifc_properties, "my_property_IsExternal")
        row = box.row()
        row.prop(ifc_properties, "my_property_LoadBearing")
        row = box.row()
        row.prop(ifc_properties, "my_property_FireRating")
        row = box.row()
        row.prop(ifc_properties, "my_property_AcousticRating")
class CUSTOM_PROPERTIES_IFC_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "Custom Properties"

    def draw(self, context):
        ifc_properties = context.scene.ifc_properties

        layout = self.layout
        box = layout.box()

        custom_collection = context.scene.custom_collection

        row = layout.row(align=True)
        row.operator("custom.collection_actions", text="Add", icon="ADD").action = "add"
        row.operator("custom.collection_actions", text="Remove Last", icon="REMOVE").action = "remove"

        for item in custom_collection.items:
            box.prop(item, "name")

        row = layout.row(align=True)    

             
class SPREADSHEET_IFC_FILE_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "Spreadsheet"

    def draw(self, context):
        ifc_properties = context.scene.ifc_properties
        layout = self.layout
       
        box_spreadsheet = layout.box()
        row = box_spreadsheet.row()
        row.prop(ifc_properties, "my_spreadsheetfile")

        row = box_spreadsheet.row()
        row.prop(ifc_properties, "ods_or_xlsx")
        box_spreadsheet.operator("export.tospreadsheet")   
class FILTER_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "Filter IFC elements"

    def draw(self, context):
        layout = self.layout

        self.layout.operator("object.filter_ifc_elements", text="Filter IFC elements", icon="FILTER")
        self.layout.operator("object.unhide_all", text="Unhide IFC elements", icon="LIGHT")
class BASE_QUANTITIES_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "BaseQuantities"

    def draw(self, context):

        ifc_properties = context.scene.ifc_properties
        layout = self.layout
        
        box = layout.box()
        row = box.row()

        row.prop(ifc_properties, "my_quantity_Length")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_Width")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_Height")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_Area")
        row = box.row()       
        row.prop(ifc_properties, "my_quantity_NetArea")
        row = box.row() 
        row.prop(ifc_properties, "my_quantity_NetSideArea")
        row = box.row() 
        row.prop(ifc_properties, "my_quantity_GrossSideArea")
        row = box.row() 
        row.prop(ifc_properties, "my_quantity_GrossArea")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_Volume")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_NetVolume")
        row = box.row()
        row.prop(ifc_properties, "my_quantity_GrossVolume")
class SAVE_SELECTION_PT_PANEL(GENERAL_panel, Panel):
    bl_parent_id = "EXAMPLE_PT_panel_1"
    bl_label = "Save Selection"

    def draw(self, context):

        ifc_properties = context.scene.ifc_properties
        layout = self.layout
        
        box = layout.box()
        row = box.row()
        row.prop(ifc_properties, "my_selectionload")
        box.operator("save.confirm_selection", text="Confirm and set UI")
        box.operator("save.save_and_load_selection")
        

classes = ( 
            GENERAL_PT_PANEL,
            GENERAL_IFC_PT_PANEL,
            COMMON_PROPERTIES_IFC_PT_PANEL,
            BASE_QUANTITIES_PT_PANEL,
            CUSTOM_PROPERTIES_IFC_PT_PANEL,
            SAVE_SELECTION_PT_PANEL,
            SPREADSHEET_IFC_FILE_PT_PANEL,
            FILTER_PT_PANEL,
            )

def register():
    for cls in classes:
        bpy.utils.register_class(cls)

def unregister():
    for cls in classes:
        bpy.utils.unregister_class(cls)