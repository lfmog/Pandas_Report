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
ws = arcpy.env.workspace = r"Mant.gdb" #--> GDB PATH
##excelOut = arcpy.GetParameterAsText(1)
excelOut = r"Borrardor" #--> TRASH PATH
features = arcpy.ListFeatureClasses()
try:
    for fc in features:
        if arcpy.Exists(fc)== True:
            def HCA(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME', 'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName,sheetname=feature_ly)
                CHCA = df.loc[:,[lista de campos]]
                df1 = CHCA.copy()
                # sort_df1 = df1.sort_values(['TRM_RML', 'PK_INICIO'],ascending=[True, True])
                # sort_df2 = sort_df1.groupby(['DUCTO','ZONA']).sum()
                # print(sort_df2)
                # sort_df1['NUM_EDIF'] = sort_df1['DISTANCIA'].sum()
                sort_df1 = df1.pivot_table(index=['DUCTO', 'ZONA', 'DISTRITO'], columns='CRITERIO',values='DISTANCIA', aggfunc='sum')
                df2 = df1.pivot_table(index=['ZONA'], columns='CRITERIO', values='DISTANCIA', aggfunc='sum')
                print(sort_df1)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                sort_df1.to_excel(writer,'Sheet1')
                writer2 = ExcelWriter(os.path.join(excelOut, feature +"_ZONA"+ ".xlsx"))
                df2.to_excel(writer2, 'Sheet2')
                writer.save()
                writer2.save()
##                print(TOTAL)
            HCA("H_AREA", "H_AREA_ly")

            def CA(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME', 'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName,sheetname=feature_ly)
                CA = df.loc[:,['FECHA', 'DUCTO', 'TRM_RML', 'ZONA','DISTRITO', 'PK_INICIO', 'PK_FIN'
                               ,'CLASE_LOC','LONG_TRM','ID', 'ENTREGA', 'FUENTE']]
                df1 = CA.copy()
                # sort_df1 = df1.sort_values(['TRM_RML', 'PK_INICIO'],ascending=[True, True])
                # sort_df2 = sort_df1.groupby(['DUCTO', 'ZONA', 'DISTRITO','CLASE_LOC']).sum()
                df2=df1.pivot_table(index = ['DUCTO','ZONA','DISTRITO','FUENTE'], columns='CLASE_LOC', values='LONG_TRM', aggfunc='sum')
                df3 = df1.pivot_table(index=['ZONA'], columns='CLASE_LOC', values='LONG_TRM', aggfunc='sum')
                print(df2)
                print(df3)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer,'Sheet1')
                writer2 = ExcelWriter(os.path.join(excelOut, feature +"_ZONA"+ ".xlsx"))
                df3.to_excel(writer2, 'Sheet2')
                writer.save()
                writer2.save()
##                print(CHCA)
            CA("C_AREA", "C_AREA_ly")

            def SE_T(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO', 'TIPO','TABLERO']]
                df1 = CA.copy()
                df2 = df1.pivot_table(index=['DUCTO', 'ZONA', 'DISTRITO'], columns=['TABLERO'], aggfunc='count')
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature +"_T"+ ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            SE_T("S_EXISTENTE", "S_EXISTENTE_ly")


            def SE_P(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO', 'TIPO', 'PEDESTAL']]
                df1 = CA.copy()
                df2 = df1.pivot_table(index=['DUCTO', 'ZONA', 'DISTRITO'], columns=['PEDESTAL'], aggfunc='count')
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature +"_P"+ ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            SE_P("S_EXISTENTE", "S_EXISTENTE_ly")

            def SE_C(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO', 'TIPO', 'CIMENTACIO']]
                df1 = CA.copy()
                df2 = df1.pivot_table(index=['DUCTO', 'ZONA', 'DISTRITO'], columns=['CIMENTACIO'], aggfunc='count')
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature +"_C"+ ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            SE_C("S_EXISTENTE", "S_EXISTENTE_ly")

            def SS(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO']]
                df1 = CA.copy()
                df2 = df1.groupby(['DUCTO', 'ZONA', 'DISTRITO'])['DUCTO', 'ZONA', 'DISTRITO'].count()
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            SS("S_SUGERIDA", "S_SUGERIDA_ly")

            def IT(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO', 'TIPO', 'TIPO_PK']]
                df1 = CA.copy()
                df2 = df1.pivot_table(index=['DUCTO', 'ZONA', 'DISTRITO'], columns=['TIPO'], aggfunc='count')
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            IT("I_TERCEROS", "I_TERCEROS_ly")

            def CA(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO']]
                df1 = CA.copy()
                df2 = df1.groupby(['DUCTO', 'ZONA', 'DISTRITO'])['DUCTO', 'ZONA', 'DISTRITO'].count()
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            CA("C_AEREO", "C_AEREO_ly")

            def CS(feature, feature_ly):
                arcpy.MakeFeatureLayer_management(feature, feature_ly)
                xls = arcpy.TableToExcel_conversion(feature_ly, (os.path.join(excelOut, feature_ly + ".xls")), 'NAME',
                                                    'DESCRIPTION')
                fileName = (os.path.join(excelOut, feature_ly + ".xls"))
                df = pd.read_excel(fileName, sheetname=feature_ly)
                CA = df.loc[:, ['DUCTO', 'ZONA', 'DISTRITO']]
                df1 = CA.copy()
                df2 = df1.groupby(['DUCTO', 'ZONA', 'DISTRITO'])['DUCTO', 'ZONA', 'DISTRITO'].count()
                print(df2)
                writer = ExcelWriter(os.path.join(excelOut, feature + ".xlsx"))
                df2.to_excel(writer, 'Sheet1')
                writer.save()
            CS("C_SUBFLUVIAL", "C_SUBFLUVIAL_ly")
        break
except:
    print(arcpy.GetMessages())
