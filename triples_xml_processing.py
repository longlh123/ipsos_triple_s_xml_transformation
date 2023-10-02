import sys
sys.path.append('c:\\Users\\long.pham\\Documents\\MDDPython')

import os
import shutil
import pandas as pd
import numpy as np
import re
import win32com.client as w32
from metadata import Metadata
import xml.etree.ElementTree as ET
from object.iSurvey import iSurvey

import collections.abc
#hyper needs the four following aliases to be done manually.
collections.Iterable = collections.abc.Iterable
collections.Mapping = collections.abc.Mapping
collections.MutableSet = collections.abc.MutableSet
collections.MutableMapping = collections.abc.MutableMapping

os.chdir("ipsos_triple_s_xml_transformation")

root = "projects"
project_name = "VN2023282TRIPLES_SYNCOPA"

xml_file = r"sources\230951.xml"
ascii_file = r"sources\230951.dat"
open_ended_file = r"sources\230951.xlsx"

isurvey = iSurvey(r'{}\{}\{}'.format(root, project_name, xml_file)) 

excel_path = "Download Survey Data for Import.xlsx"
source_mdd_file = r"template\TemplateProject.mdd"
current_mdd_file = "{}\{}\{}.mdd".format(root, project_name, project_name)
source_dms_file = r"dms\OutputDDFFile.dms"

if not os.path.exists(current_mdd_file):
    shutil.copy(source_mdd_file, current_mdd_file)

mdd_source = Metadata(mdd_file=current_mdd_file, dms_file=source_dms_file)

for variable_name, variable in isurvey["variables"].items():
    if "syntax" in list(variable.keys()):   
        mdd_source.addScript(variable_name, variable["syntax"])

mdd_source.runDMS()

df_datasource = pd.read_csv(r'{}\{}\{}'.format(root, project_name, ascii_file), delimiter='\t', header=None)
df_oe_datasource = pd.read_excel(r'{}\{}\{}'.format(root, project_name, open_ended_file), engine="openpyxl")
df_oe_datasource.set_index(['record'], inplace=True)

adoConn = w32.Dispatch('ADODB.Connection')
conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(mdd_source.mdd_file.replace('.mdd', '_EXPORT.ddf'), mdd_source.mdd_file.replace('.mdd', '_EXPORT.mdd'))
adoConn.Open(conn)

sql_delete = "DELETE FROM VDATA"
adoConn.Execute(sql_delete)

for i, row in df_datasource[list(df_datasource.columns)].iterrows():
    try:
        start = isurvey["variables"]['record']['position']['start']
        length = isurvey["variables"]['record']['position']['finish']
        record_id = re.sub(pattern="\s", repl="", string=row[0][start:length])

        sql_insert = "INSERT INTO VDATA(record) VALUES(%s)" % (record_id)
        adoConn.Execute(sql_insert)
        
        c = list()
        v = list()

        for variable_name, variable in isurvey["variables"].items():
            if variable_name not in ['date','record']:
                start = isurvey["variables"][variable_name]['position']['start']
                length = isurvey["variables"][variable_name]['position']['finish']
                
                if len(re.sub(pattern="\s", repl="", string=row[0][start:length])) > 0:
                    if variable_name == "Q1":
                        s = ""
                    
                    c.append(variable_name)

                    match variable['type']:
                        case 'quantity':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append(value)
                        case 'character':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("'{}'".format(value))
                        case 'single':
                            value = re.sub(pattern="\s", repl="", string=row[0][start:length])
                            v.append("{_%s}" % (value))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))
                        case 'multiple':
                            value = row[0][start:length]
                            
                            arr = list()

                            for i in range(len(value)):
                                if len(re.sub(pattern="\s", repl="", string=value[i])) > 0:
                                    if int(value[i]) == 1:
                                        arr.append("_{}".format(list(variable['values'].keys())[int(i)]))
                            
                            v.append("{%s}" % (",".join(arr)))

                            for code, helperfield in variable['helperfields'].items():
                                if int(record_id) in list(df_oe_datasource.index):
                                    if pd.notnull(df_oe_datasource.loc[int(record_id), helperfield['name']]):
                                        c.append("{}.{}".format(variable_name, helperfield['name']))
                                        v.append("'{}'".format(df_oe_datasource.loc[int(record_id), helperfield['name']]))

        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE record = {}".format(record_id)
        adoConn.Execute(sql_update)    
    except Exception as ex:
        print(sql_insert, ex, sep="-")
        sys.exit(1)



