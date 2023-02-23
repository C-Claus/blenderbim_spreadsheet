import bpy
from . import prop

import collections
from collections import defaultdict

import ifcopenshell
import pandas as pd

class ConstructDataFrame:
    def __init__(self, context):
        print ('hallo uit Construct dataframe class')

        ifc_dictionary = defaultdict(list)
        ifc_properties = context.scene.ifc_properties

        #ifc_file = ifcopenshell.open(IfcStore.path)
        ifc_file = ifcopenshell.open("C:\\Algemeen\\07_ifcopenshell\\00_ifc\\02_ifc_library\\IFC Schependomlaan.ifc")
        products = ifc_file.by_type('IfcProduct')

        for product in products:
            ifc_dictionary[prop.prop_globalid].append(str(product.GlobalId))

            if ifc_properties.my_ifcproduct:
                ifc_dictionary[prop.prop_ifcproduct].append(str(product.is_a()))
            
            if ifc_properties.my_ifcproductname:
                ifc_dictionary[prop.prop_ifcproductname].append(str(product.Name))

            if ifc_properties.my_ifcproducttypename:
                ifc_dictionary[prop.prop_ifcproducttypename].append(self.get_ifc_type(context, ifc_product=product)[0])

            if ifc_properties.my_ifcbuildingstorey:
                ifc_dictionary[prop.prop_ifcbuildingstorey].append(self.get_ifc_building_storey(context, ifc_product=product)[0])

            if ifc_properties.my_ifcclassification:
                ifc_dictionary[prop.prop_classification].append(self.get_ifc_classification_item_and_reference(context, ifc_product=product)[0])


        df = pd.DataFrame(ifc_dictionary)
        self.df = df

        print (self.df)

    def get_ifc_type(self, context, ifc_product):
    
        ifc_type_list = []
        
        if ifc_product: 
            ifcproduct_type = ifcopenshell.util.element.get_type(ifc_product)
    
            if ifcproduct_type:
                ifc_type_list.append(ifcproduct_type.Name)
        
        if not ifc_type_list:
            ifc_type_list.append(None)

        return ifc_type_list

    def get_ifc_building_storey(self, context,ifc_product):

        building_storey_list = []
            
        spatial_container = ifcopenshell.util.element.get_container(ifc_product)
        
        if spatial_container:    
            building_storey_list.append(spatial_container.Name)
            
        if not building_storey_list:
            building_storey_list.append(None)
            
        return building_storey_list

    def get_ifc_classification_item_and_reference(self, context, ifc_product):
    
        classification_list = []

        # Elements may have multiple classification references assigned
        references = ifcopenshell.util.classification.get_references(ifc_product)
        
        if ifc_product:
            for reference in references:
                system = ifcopenshell.util.classification.get_classification(reference)
                classification_list.append(str(system.Name) + ' | ' + str(reference[1]) +  ' | ' + str(reference[2]))
            
        if not classification_list:
            classification_list.append(None)  
            
        return classification_list



class ExportToSpreadSheet(bpy.types.Operator):
    bl_idname = "export.tospreadsheet"
    bl_label = "Export to Spreadsheet"

    def execute(self, context):

        ifc_properties = context.scene.ifc_properties
        construct_data_frame = ConstructDataFrame(context)

        #this adds nothing, selection is allreadye happing dataframes

        #if ifc_properties.my_ifcproduct:
        #    print ('prop.prop_ifcproduct', prop.prop_ifcproduct)
        #    print ('ifc_properties.my_ifcproduct', ifc_properties.my_ifcproduct)

        #if ifc_properties.my_ifcbuildingstorey:
        #    print ('prop.ifc_buildingstorey', prop.prop_ifcbuildingstorey)
        #    print ('ifc_properties.ifc_buildingstorey', ifc_properties.my_ifcbuildingstorey)

        return {'FINISHED'}

def register():
    bpy.utils.register_class(ExportToSpreadSheet)

def unregister():
    bpy.utils.unregister_class(ExportToSpreadSheet)





        


"""
import bpy
import pandas
import collections
from collections import defaultdict

import ifcopenshell
import ifcopenshell.api
import ifcopenshell.util.element
import ifcopenshell.util.classification
#import ifcopenshell.util.material

import blenderbim
#import blenderbim.bim.import_ifc
#from blenderbim.bim.ifc import IfcStore

import pandas as pd

ifc_file_location = "C:\\Algemeen\\07_ifcopenshell\\00_ifc\\02_ifc_library\\IFC Schependomlaan.ifc"
#ifc_file_location = "C:\\Algemeen\\07_ifcopenshell\\00_ifc\\02_ifc_library\\IFC4 demo.ifc"
ifc_file = ifcopenshell.open(ifc_file_location)

products = ifc_file.by_type('IfcProduct')

def get_ifc_type(self, context, ifc_product):
    
    ifc_type_list = []
    
    if ifc_product: 
        ifcproduct_type = ifcopenshell.util.element.get_type(ifc_product)
 
        if ifcproduct_type:
            ifc_type_list.append(ifcproduct_type.Name)
    
    if not ifc_type_list:
        ifc_type_list.append(None)

    return ifc_type_list


def get_ifc_building_storey(ifc_product):

    building_storey_list = []
        
    spatial_container = ifcopenshell.util.element.get_container(ifc_product)
    
    if spatial_container:    
        building_storey_list.append(spatial_container.Name)
        
    if not building_storey_list:
        building_storey_list.append(None)
         
    return building_storey_list 


def get_ifc_classification_item_and_reference(ifc_product):
    
    classification_list = []

    # Elements may have multiple classification references assigned
    references = ifcopenshell.util.classification.get_references(ifc_product)
    
    if ifc_product:
        for reference in references:
            system = ifcopenshell.util.classification.get_classification(reference)
            classification_list.append(str(system.Name) + ' | ' + str(reference[1]) +  ' | ' + str(reference[2]))
        
    if not classification_list:
        classification_list.append(None)  
          
    return (classification_list)
    
    
def get_ifc_materials(ifc_product):
    
    material_list = []
     
    if ifc_product:
        ifc_material = ifcopenshell.util.element.get_material(ifc_product)
        if ifc_material:
            
            if ifc_material.is_a('IfcMaterial'):
                material_list.append(ifc_material.Name)
            
            if ifc_material.is_a('IfcMaterialList'):
                for materials in ifc_material.Materials:
                   material_list.append(materials.Name)
            
            if ifc_material.is_a('IfcMaterialConstituentSet'):
                for material_constituents in ifc_material.MaterialConstituents:
                    material_list.append(material_constituents.Material.Name)
            
            if ifc_material.is_a('IfcMaterialLayerSetUsage'):
                for material_layer in ifc_material.ForLayerSet.MaterialLayers:
                    material_list.append(material_layer.Material.Name)
                
            if ifc_material.is_a('IfcMaterialProfileSetUsage'):
                for material_profile in (ifc_material.ForProfileSet.MaterialProfiles):
                    material_list.append(material_profile.Material.Name)
    
    if not material_list:
        material_list.append(None)
        
    return material_list


def get_ifc_properties_and_quantities(ifc_product, ifc_propertyset_name, ifc_property_name):
    
    ifc_property_list = []
    
    if ifc_product:
        ifc_property_list.append(ifcopenshell.util.element.get_pset(ifc_product, ifc_propertyset_name,ifc_property_name))
        
    if not ifc_property_list:
        ifc_property_list.append(None)
        
    return ifc_property_list






class ConstructPandasDataFrame:
    
    def __init__(self, context):
        
        ifc_dictionary = defaultdict(list)
        
        for product in products:
            if product:
                ifc_pset_common = 'Pset_' +  (str(product.is_a()).replace('Ifc','')) + 'Common'
                

                ifc_dictionary[global_id].append(str(product.GlobalId))
                ifc_dictionary[ifc_product_is].append(str(product.is_a()))
                ifc_dictionary[ifc_name].append(str(product.Name))
                ifc_dictionary[ifc_type].append(str(get_ifc_type(self, context, ifc_product=product)[0]))
                
                
                ifc_dictionary[ifc_buildingstorey].append(get_ifc_building_storey(ifc_product=product)[0])
                ifc_dictionary[ifc_classification].append(get_ifc_classification_item_and_reference(ifc_product=product)[0])
                
                
                ifc_dictionary[ifc_material].append(get_ifc_materials(ifc_product=product)[0])
               
                
                ifc_dictionary[is_external].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                     ifc_propertyset_name=ifc_pset_common,
                                                                                     ifc_property_name=is_external)[0])
                                                                                
                ifc_dictionary[load_bearing].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                      ifc_propertyset_name=ifc_pset_common,
                                                                                      ifc_property_name=load_bearing)[0])
                                                                                
                ifc_dictionary[fire_rating].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                     ifc_propertyset_name=ifc_pset_common,
                                                                                     ifc_property_name=fire_rating)[0])
                                                                                
                                                                                
                ifc_dictionary[area].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                ifc_propertyset_name=base_quantities,
                                                                                ifc_property_name=area)[0])
                                                                                
                ifc_dictionary[net_area].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                   ifc_propertyset_name=base_quantities,
                                                                                   ifc_property_name=net_area)[0])
                                                                                
                ifc_dictionary[netside_area].append(get_ifc_properties_and_quantities(ifc_product=product,
                                                                                       ifc_propertyset_name=base_quantities,
                                                                                       ifc_property_name=netside_area)[0])
                                                                                       
                
                                                                                       
               
            
        #print (ifc_dictionary)
        
        df = pd.DataFrame(ifc_dictionary)
        self.df = df
            
            
        #print (df)

   
        
class WriteToXLSXSpreadSheet(bpy.types.Operator):
    
    bl_idname = "object.write_to_xlsx_spreadsheet"
    bl_label = "Simple Operator"

    def execute(self, context):
        print (' execute from write to xlsx')
        
        #data_frame = ConstructPandasDataFrame()            
        #data_frame.construct_dataframe()  
        
        #df_builder = ConstructDataFrame(context)
        #df = df_builder.create_dataframe()
        
        
        
        #data_frame = ConstructPandasDataFrame(context)
        #data_frame.construct_dataframe()
        
        #print (data_frame)
        
        return {'FINISHED'}

class WriteToODS(bpy.types.Operator):

    def execute():
        print (' execute')

class FilterIFCElements(bpy.types.Operator):

    def execute():
        print (' execute')

class UnhideIFCElements(bpy.types.Operator):

    def execute():
        print (' execute')

class GetCustomCollection(bpy.types.Operator):

    def execute():
        print (' execute')
        


#overal self context tusssen plaatsen
bpy.ops.object.write_to_xlsx_spreadsheet('EXEC_DEFAULT')

#write to xlsx
#write to ods
#start with ui, vul ui met data uit ifc 

def register():
    bpy.utils.register_class(WriteToXLSXSpreadSheet)

# Unregister the operator
def unregister():
    bpy.utils.unregister_class(WriteToXLSXSpreadSheet)
    
if __name__ == "__main__":
    register()
"""