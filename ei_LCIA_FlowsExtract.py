# -*- coding: utf-8 -*-
"""

Last modified June 2021
@author: pauliuk

see: https://github.com/IndEcol/openLCA_ecoinvent_Material_Footprint_LCIA

"""

# Script ei_LCIA_FlowsExtract.py

# Import required libraries:
#%%

import os
import pandas as pd    
import uuid
import json

import mf_Paths

#################
#     Functions #
#################

None

#################
#     MAIN      #
#################
"""
Set configuration data
"""    

ei_version_string = '_ei_3_7_1'
#ei_version_string = '_ei_3_8'

ScriptConfig = {}
ScriptConfig['Current_UUID'] = str(uuid.uuid4())

###################################################################################
#   Import data                                                                   #
###################################################################################   
#%%
if ei_version_string == '_ei_3_7_1':
    tp = mf_Paths.data_path_ei371
if ei_version_string == '_ei_3_8':
    tp = mf_Paths.data_path_ei38
    
# sample LCIA impact category
testfile_in = os.path.join(tp,'lcia_categories','0e54d54c-fdcb-35d7-aca2-cc3586eee3e7.json')
with open(testfile_in, "r") as read_file:
    data_lcia_sample = json.load(read_file)
    
#sample flow:    
testfile_in = os.path.join(tp,'flows','0ae80e4e-2e46-5292-a0d7-7526bebc603b.json')
with open(testfile_in, "r") as read_file:
    data_flow_sample = json.load(read_file)    
data_flow_sample['name']
data_flow_sample['flowType']    
data_flow_sample['@id']       
data_flow_sample['category']['categoryPath'][0]
data_flow_sample['category']['categoryPath'][1]
data_flow_sample['flowProperties'][0]['flowProperty']['name']
data_flow_sample['flowProperties'][0]['flowProperty']['name']

#xml file with characterisation factors:
# Does not work: throws formatting error 'not well formed'.
# Will rather extract flows to Excel and then make a manual match to all 10 categories.    
#char_factor_xml_in = os.path.join(mcPaths.data_path,'resources-08-00061-s001.xml')
#xmldict = xmltodict.parse(char_factor_xml_in)   
    
###################################################################################
#   Extract and sort data                                                         #
################################################################################### 
#%%

#loop over all flows:
flowdirectory = os.fsencode(os.path.join(tp,'flows'))
flow_df = pd.DataFrame(columns = ['Name', 'Type','id','category1','category2','category3','flowproperty_name','flowproperty_id','flowproperty_categorypath',])
m=0
for file in os.listdir(flowdirectory):
      filename = os.fsdecode(file)
      print(filename)
      with open(os.path.join(tp,'flows',filename), "r") as read_file:
          data_flow_sample = json.load(read_file)     
      try:
          flow_df.loc[m] = [data_flow_sample['name'],data_flow_sample['flowType'],data_flow_sample['@id'],data_flow_sample['category']['categoryPath'][0],data_flow_sample['category']['categoryPath'][1],data_flow_sample['category']['categoryPath'][2],data_flow_sample['flowProperties'][0]['flowProperty']['name'],data_flow_sample['flowProperties'][0]['flowProperty']['@id'],data_flow_sample['flowProperties'][0]['flowProperty']['categoryPath'][0]]  
      except: 
          flow_df.loc[m] = [data_flow_sample['name'],data_flow_sample['flowType'],data_flow_sample['@id'],data_flow_sample['category']['categoryPath'][0],data_flow_sample['category']['categoryPath'][1],'',data_flow_sample['flowProperties'][0]['flowProperty']['name'],data_flow_sample['flowProperties'][0]['flowProperty']['@id'],data_flow_sample['flowProperties'][0]['flowProperty']['categoryPath'][0]]  
      m +=1

flow_df.to_excel(os.path.join(mf_Paths.result_path,'elementary_flows' + ei_version_string + '.xlsx'))
    
#%%
#
#
# The End
#
#
