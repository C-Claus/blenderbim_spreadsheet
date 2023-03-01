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

quantity_list = [
                 prop_length,
                 prop_width,
                 prop_height,
                 prop_area,
                 prop_netarea,
                 prop_netsidearea,
                 prop_grossarea,
                 prop_volume,
                 prop_netvolume,
                 prop_grossvolume]

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

    my_isexternal:              bpy.props.BoolProperty(name=prop_isexternal,        default=True)
    my_loadbearing:             bpy.props.BoolProperty(name=prop_loadbearing,       default=True)
    my_firerating:              bpy.props.BoolProperty(name=prop_firerating,        default=True)
    my_acousticrating:          bpy.props.BoolProperty(name=prop_acousticrating,    default=False)


    my_length:                  bpy.props.BoolProperty(name=prop_length,            default=False)
    my_width:                   bpy.props.BoolProperty(name=prop_width,             default=True)
    my_height:                  bpy.props.BoolProperty(name=prop_height,            default=True)

    my_area:                    bpy.props.BoolProperty(name=prop_area,              default=False)
    my_netarea:                 bpy.props.BoolProperty(name=prop_netarea,           default=False)
    my_netsidearea:             bpy.props.BoolProperty(name=prop_netsidearea,       default=False)
    my_grossarea:               bpy.props.BoolProperty(name=prop_grossarea,         default=False)
    my_grosssidearea:           bpy.props.BoolProperty(name=prop_grosssidearea,     default=False)

    my_volume:                  bpy.props.BoolProperty(name=prop_volume,            default=True)
    my_netvolume:               bpy.props.BoolProperty(name=prop_netvolume,         default=True)
    my_grossvolume:             bpy.props.BoolProperty(name=prop_grossvolume,       default=True)
  
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
