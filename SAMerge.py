
import docx, os, pandas as pd, numpy as np, pickle as pk, openpyxl as opxl, main_pandas
from datetime import datetime as dt

arrPath = main_pandas.arrPath
strUnit = main_pandas.strUnit
strSAMinput = main_pandas.strSAMinput
strSALinput = main_pandas.strSALinput
strSALtab = main_pandas.strSALtab
strParaFile = main_pandas.strParaFile
strSAMtp = main_pandas.strSAMtp
strTagLookup = main_pandas.strTagLookup
arrSALheader = main_pandas.arrSALheader
strSAMOutput = main_pandas.strSAMoutput


# Load Input
with open(os.path.join(*arrPath, strSAMinput), "rb") as pkLoader:
    dfSAD = pk.load(pkLoader)
dfSAL = pd.read_excel(os.path.join(*arrPath, strSALinput), strSALtab)
dfSAMtemp = pd.read_excel(os.path.join('.', strSAMtp),'MailMerge')

xlwbSAM = opxl.load_workbook(strSAMtp, read_only=False, keep_vba=True)
xlwsSAM = xlwbSAM['MailMerge']


# Clean up - SAM
dfSAM_header = dfSAMtemp.drop(columns=['Unnamed: 0',strTagLookup,'SIL_SIF','meet_SIF','rec1_SIF','rec2_SIF','rec3_SIF']).columns.values


# Clean up - SAL
dfSAL = dfSAL.drop(arrSALheader, axis=0)
dfSAL = dfSAL.set_index(strTagLookup)


# Clean up - SAD
dfSAD = dfSAD.set_index(strTagLookup)


# Merge SAD and SAL
dfSAM = pd.concat([dfSAD,dfSAL],axis=1,ignore_index=False,sort=False)

dfSAM = dfSAM.reindex(dfSAM_header, axis=1)
dfSAM.index.name = strTagLookup


# Write DataFrame to SAMerge Template
def write_DF_to_Excelws(xlwsObj, dfObj, numRowOffSet=0, numColOffSet=0):
    lenDFRow, lenDFCol = dfObj.shape
    lenDFCol += 1   # counting index column
    lenWSRow = lenDFRow + numRowOffSet
    lenWSCol = lenDFCol + numColOffSet
    xlwsObj.cell(row=lenWSRow, column=lenWSCol).value = None   # allocate memory
    iterWSrows = xlwsObj.iter_rows()
    iterDFrows = dfObj.itertuples(name=None)
    for _ in range(numRowOffSet):
        _ = next(iterWSrows)
    for seqRow in range(lenDFRow):
        rowWS = next(iterWSrows)
        rowDF = next(iterDFrows)
        for seqCol in range(lenDFCol):
            rowWS[seqCol + numColOffSet].value = rowDF[seqCol]

write_DF_to_Excelws(xlwsSAM, dfSAM, numRowOffSet=3, numColOffSet=1)


# Export SAM
strTimeNow = dt.now().strftime("%d%b%y")
xlwbSAM.save(f'{os.path.join(*arrPath, "")}{strSAMOutput}-{strTimeNow}.xlsm')

