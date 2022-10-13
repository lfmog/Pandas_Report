#!/usr/bin/env python
# -*- coding: utf-8 -*-
import arcpy,os, string,re, math
from arcpy import env
import requests
import time
import argparse
import unicodedata
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
##sys.stdout.reconfigure(encoding='utf-8')
arcpy.env.overwriteOutput = True;
import pandas as pd
from pandas import ExcelWriter
arcpy.env.overwriteOutput = True;
##ws = arcpy.env.workspace = arcpy.GetParameterAsText(0)
ws = arcpy.env.workspace = r"D.gdb"
##excelOut = arcpy.GetParameterAsText(1)
excelOut = r"D:\Process\PRUEBAS_PYTHON"
features = arcpy.ListFeatureClasses()
try:
    for fc in features:
        if arcpy.Exists(fc)== True:
            def SP(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME', 'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                df = df.assign(PK_FIN="")
                CA = df.loc[:, ['OBJECTID', 'DUCTO','ZONA', 'DISTRITO', 'TRM_RML', 'TIPO_PK', 'PK', 'PK_FIN', 'TP_DUCTO', 'SUB_CLASE']]
                # CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO', 'TRM_RML', 'SUB_CLASE', 'TIPO_PK', 'PK']]
                df1 = CA.copy()
                sort_df1 = df1.sort_values(['DUCTO', 'TRM_RML','TP_DUCTO','PK'], ascending=[True, True, True, True])
                # sort_df1 = df1.sort_values(['OBJECTID'], ascending=[True])
                dfOut = pd.DataFrame(columns=sort_df1.columns)
                inicios = 0
                final = 0
                cnt = 0
                print("Geotecnia: Calculando INICIO-> FIN ")
                for i in range(0, sort_df1.shape[0]):
                    if cnt == 1:
                        final = i
                        sort_df1.iloc[inicios, sort_df1.columns.get_loc('PK')], sort_df1.iloc[final, sort_df1.columns.get_loc('PK')]
                        # print (sort_df2)
                        sort_df1.iloc[inicios, sort_df1.columns.get_loc('TIPO_PK')], sort_df1.iloc[final, sort_df1.columns.get_loc('TIPO_PK')]
                        # print (sort_df3)
                        sort_df1.iloc[inicios, sort_df1.columns.get_loc('PK_FIN')] = sort_df1.iloc[final, sort_df1.columns.get_loc('PK')]
                        print (sort_df1)
                        dfOut = dfOut.append(sort_df1.iloc[inicios], ignore_index=True)
                        print (dfOut)
                        cnt = 0
                    else:
                        inicios = i
                        cnt += 1
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                dfOut.to_excel(writer, 'Sheet1')
                writer.save()
            SP("SIN_PASO", "S_PASO_ly")
            def DUCTO(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME', 'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                df = df.assign(PK_FIN="")
                CA = df.loc[:, ['ENTREGA', 'DUCTO', 'TRM_RML', 'TP_DUCTO']]
                df1 = CA.copy()
                df2 = df1.groupby(['ENTREGA', 'DUCTO', 'TRM_RML', 'TP_DUCTO']).count()
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            DUCTO("DUCTO", "DUCTO_ly")
        break
except:
    print(arcpy.GetMessages())
