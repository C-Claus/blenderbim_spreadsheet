import bpy
from bpy.types import Scene
from bpy.props import BoolProperty, StringProperty


prop_ifcproduct         = 'IfcProduct'
prop_ifcbuildingstorey  = 'IfcBuildingStorey'
prop_ifcproductname     = 'Name'
prop_ifcproducttypename = 'Type'
prop_isexternal         = 'IsExternal'
prop_loadbearing        = 'LoadBearing'
prop_firerating         = 'FireRating'
prop_length             = 'Length'
prop_width              = 'Width'
prop_area               = 'Area'
prop_netarea            = 'NetArea'
prop_netsidearea        = 'NetSideArea'

class IFCProperties(bpy.types.PropertyGroup):
    my_ifcproduct: bpy.props.BoolProperty(name=prop_ifcproduct,description="Export IfcProduct",default=True)
    my_ifcbuildingstorey: bpy.props.BoolProperty(name=prop_ifcbuildingstorey,description="Export IfcBuildingStorey",default = True)
  

def register():
    bpy.utils.register_class(IFCProperties)
    bpy.types.Scene.ifc_properties = bpy.props.PointerProperty(type=IFCProperties)

def unregister():
    bpy.utils.unregister_class(IFCProperties)
    del bpy.types.Scene.ifc_properties


""" 
def register():

    #my_ifcproduct: bpy.props.BoolProperty(name="IfcProduct",description="Export IfcProduct",default=True)

    Scene.property_ifcproduct = BoolProperty(name="IfcProduct",description="Export IfcProduct",default=True)

    #Scene.my_property = BoolProperty(default=True)
    #Scene.another_property = BoolProperty(default=True)
    #Scene.yet_another_property = BoolProperty(default=False)
    #Scene.and_yet_another_property = BoolProperty(default=False)

def unregister():

    del Scene.property_ifcproduct

   

    #del Scene.my_property
    #del Scene.another_property
    #del Scene.yet_another_property
    #del Scene.and_yet_another_property
    

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
       
class BlenderBIMSpreadSheetProperties(bpy.types.PropertyGroup):
    ###############################################
    ################# General #####################
    ############################################### 
    my_ifcproduct: bpy.props.BoolProperty(name="IfcProduct",description="Export IfcProduct",default=True)
    my_ifcbuildingstorey: bpy.props.BoolProperty(name="IfcBuildingStorey",description="Export IfcBuildingStorey",default = True)     
    my_ifcproduct_name: bpy.props.BoolProperty(name="Name",description="Export IfcProduct Name",default = True)
    my_type: bpy.props.BoolProperty(name="Type",description="Export IfcObjectType Name",default = True)
    my_ifcclassification: bpy.props.BoolProperty(name="IfcClassification",description="Export Classification",default = True)
    my_ifcmaterial: bpy.props.BoolProperty(name="IfcMaterial",description="Export Materials",default = True)
     
    ###############################################
    ############ Common Properties ################
    ###############################################
    my_isexternal: bpy.props.BoolProperty(name="IsExternal",description="Export IsExternal",default = True)
    my_loadbearing: bpy.props.BoolProperty(name="LoadBearing",description="Export LoadBearing",default = True)
    my_firerating: bpy.props.BoolProperty(name="FireRating",description="Export FireRating",default = True)
    
    ###############################################
    ############# BaseQuantities ##################
    ###############################################
    my_length: bpy.props.BoolProperty(name="Length",description="Export Length from BaseQuantities",default = True)  
    my_width: bpy.props.BoolProperty(name="Width",description="Export Width from BaseQuantities",default = True)   
    my_height: bpy.props.BoolProperty(name="Height",description="Export Height from BaseQuantities",default = True) 
   
    my_area: bpy.props.BoolProperty(name="Area",description="Gets each possible defintion of Area",default = True)  
    my_netarea: bpy.props.BoolProperty(name="NetArea",description="Export NetArea from BaseQuantities",default = True)
    my_netsidearea: bpy.props.BoolProperty(name="NetSideArea",description="Export NetSideArea from BaseQuantities",default = True)
    
    my_volume: bpy.props.BoolProperty(name="Volume",description="Export Volume from BaseQuantities",default = True) 
    my_netvolume: bpy.props.BoolProperty(name="NetVolume",description="Export NetVolume from BaseQuantities",default = True)
    
    my_perimeter: bpy.props.BoolProperty(name="Perimeter",description="Export Perimeter from BaseQuantities",default = True)      
  
    ###############################################
    ####### Spreadsheet Properties ################
    ###############################################
    my_workbook: bpy.props.StringProperty(name="my_workbook")
    my_xlsx_file: bpy.props.StringProperty(name="my_xlsx_file")
    my_ods_file: bpy.props.StringProperty(name="my_ods_file")
    
    my_file_path: bpy.props.StringProperty(name="Spreadsheet",
                                        description="your .ods or .xlsx file",
                                        default="",
                                        maxlen=1024,
                                        subtype="FILE_PATH")

  
class CustomItem(bpy.types.PropertyGroup):
    name: bpy.props.StringProperty(name="Property",description="Use the PropertySet name and Property name divided by a .",default = "[PropertySet.Property]") 

class CustomCollection(bpy.types.PropertyGroup):
    items: bpy.props.CollectionProperty(type=CustomItem)
"""