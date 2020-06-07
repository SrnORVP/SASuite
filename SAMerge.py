
import docx, os, pandas as pd, numpy as np, pickle as pk, openpyxl as opxl, main_pandas, SACusFun as SACF
from datetime import datetime as dt

arrPath = main_pandas.arrPath
strUnit = main_pandas.strUnit
strSAMinput = main_pandas.strSAMinput
strSALinput = main_pandas.strSALinput
strSALtab = main_pandas.strSALtab
strSASparam = main_pandas.strSASparam
strSAMtp = main_pandas.strSAMtmple
strTagLookup = main_pandas.strTagLookup
intSALheader = main_pandas.intSALheader
strSAMOutput = main_pandas.strSAMoutput


# Load Input
with open(os.path.join(*arrPath, strSAMinput), "rb") as pkLoader:
    dfSAD = pk.load(pkLoader)
dfSAL = pd.read_excel(os.path.join(*arrPath, strSALinput), strSALtab)
dfSAMtemp = pd.read_excel(os.path.join('.', strSAMtp),'MailMerge')

xlwbSAM = opxl.load_workbook(strSAMtp, read_only=False, keep_vba=True)
xlwsSAM = xlwbSAM['MailMerge']


# Clean up - SAM
arrHeader_SAM = dfSAMtemp.drop(columns=[strTagLookup, 'SIL_SIF', 'meet_SIF', 'rec1_SIF', 'rec2_SIF', 'rec3_SIF']).columns.to_list()
for intSeq in range(len(arrHeader_SAM)-1,-1,-1):
    if 'Unnamed' in arrHeader_SAM[intSeq]:
        del arrHeader_SAM[intSeq]

# Clean up - SAL
dfSAL = dfSAL.drop(intSALheader, axis=0)
dfSAL = dfSAL.set_index(strTagLookup)

# Clean up - SAD
dfSAD = dfSAD.set_index(strTagLookup)


# Merge SAD and SAL
dfSAM = pd.concat([dfSAD,dfSAL],axis=1,ignore_index=False,sort=False)
dfSAM = dfSAM.reindex(arrHeader_SAM, axis=1)
dfSAM.index.name = strTagLookup


# Write and Export SAM
SACF.write_pdDF_to_opxlWS(xlwsSAM, dfSAM, numRowOffSet=3, numColOffSet=1,isIndexWrite=True)
strTimeNow = dt.now().strftime("%d%b%y")
xlwbSAM.save(f'{os.path.join(*arrPath, "")}{strSAMOutput}-{strTimeNow}.xlsm')

