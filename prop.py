import bpy
from bpy.types import Scene
from bpy.props import BoolProperty, StringProperty

prop_globalid           = 'GlobalId'
prop_ifcproduct         = 'IfcProduct'
prop_ifcbuildingstorey  = 'IfcBuildingStorey'
prop_ifcproductname     = 'Name'
prop_ifcproducttypename = 'Type'
prop_classification     = 'Classification(s)'
prop_materials          = 'Material(s)'

prop_psetcommon         = 'Pset_Common'

prop_isexternal         = 'IsExternal'
prop_loadbearing        = 'LoadBearing'
prop_firerating         = 'FireRating'
prop_acousticrating     = 'AcousticRating'

prop_basequantities     = 'BaseQuantities'
prop_length             = 'Length'
prop_width              = 'Width'
prop_height             = 'Height'
prop_area               = 'Area'
prop_netarea            = 'NetArea'
prop_netsidearea        = 'NetSideArea'
prop_grossarea          = 'GrossArea'
prop_grosssidearea      = 'GrossSideArea'
prop_volume             = 'Volume'
prop_netvolume          = 'NetVolume'
prop_grossvolume        = 'GrossVolume'


prop_spreadsheetfile    = 'Spreadsheet'
prop_selectionfile      = 'SelectionFile'

class IFCProperties(bpy.types.PropertyGroup):
    my_selectionload:           bpy.props.StringProperty(   name="Load Selection",
                                                            description="Load your previous saved selections",
                                                            default="",
                                                            maxlen=1024,
                                                            subtype="FILE_PATH")
    

    
    my_ifcproduct:              bpy.props.BoolProperty(name=prop_ifcproduct,        default=True)
    my_ifcproductname:          bpy.props.BoolProperty(name=prop_ifcproductname,    default=True)
    my_ifcproducttypename:      bpy.props.BoolProperty(name=prop_ifcproducttypename,default=True)
    my_ifcbuildingstorey:       bpy.props.BoolProperty(name=prop_ifcbuildingstorey, default=True)
    my_ifcclassification:       bpy.props.BoolProperty(name=prop_classification,    default=True)
    my_ifcmaterial:             bpy.props.BoolProperty(name=prop_materials,         default=True)
    my_ifcpsetcommon:           bpy.props.BoolProperty(name=prop_psetcommon,        default=True)

    my_property_IsExternal:     bpy.props.BoolProperty(name=prop_isexternal,        default=False)
    my_property_LoadBearing:    bpy.props.BoolProperty(name=prop_loadbearing,       default=False)
    my_property_FireRating:     bpy.props.BoolProperty(name=prop_firerating,        default=False)
    my_property_AcousticRating: bpy.props.BoolProperty(name=prop_acousticrating,    default=False)


    my_quantity_Length:         bpy.props.BoolProperty(name=prop_length,            default=False)
    my_quantity_Width:          bpy.props.BoolProperty(name=prop_width,             default=False)
    my_quantity_Height:         bpy.props.BoolProperty(name=prop_height,            default=False)
    my_quantity_Area:           bpy.props.BoolProperty(name=prop_area,              default=False)
    my_quantity_NetArea:        bpy.props.BoolProperty(name=prop_netarea,           default=False)
    my_quantity_NetSideArea:    bpy.props.BoolProperty(name=prop_netsidearea,       default=False)
    my_quantity_GrossArea:      bpy.props.BoolProperty(name=prop_grossarea,         default=False)
    my_quantity_GrossSideArea:  bpy.props.BoolProperty(name=prop_grosssidearea,     default=False)
    my_quantity_Volume:         bpy.props.BoolProperty(name=prop_volume,            default=False)
    my_quantity_NetVolume:      bpy.props.BoolProperty(name=prop_netvolume,         default=False)
    my_quantity_GrossVolume:    bpy.props.BoolProperty(name=prop_grossvolume,       default=False)
  
    my_spreadsheetfile:         bpy.props.StringProperty(   name="Spreadsheet",
                                                            description="your .ods or .xlsx file",
                                                            default="",
                                                            maxlen=1024,
                                                            subtype="FILE_PATH")

    ods_or_xlsx:    bpy.props.EnumProperty(
                                name="File format",
                                items=[
                                    ("ODS", ".ods", "ods"),
                                    ("XLSX", ".xlsx", "xlsx"),
                                ],
                                default="ODS",
                               
                                )
    
class CustomItem(bpy.types.PropertyGroup):
    name: bpy.props.StringProperty(name         ="Property",
                                   description  ="Use the PropertySet name and Property name divided by a .",
                                   default      ="[PropertySet.Property]") 
class CustomCollection(bpy.types.PropertyGroup):
    items: bpy.props.CollectionProperty(type=CustomItem) 
    
def register():
    bpy.utils.register_class(IFCProperties)
    bpy.utils.register_class(CustomItem)
    bpy.utils.register_class(CustomCollection)
    bpy.types.Scene.ifc_properties = bpy.props.PointerProperty(type=IFCProperties)
    bpy.types.Scene.custom_collection = bpy.props.PointerProperty(type=CustomCollection)

def unregister():
    bpy.utils.unregister_class(IFCProperties)
    bpy.utils.unregister_class(CustomItem)
    bpy.utils.unregister_class(CustomCollection)
    del bpy.types.Scene.ifc_properties
