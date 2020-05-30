
import pandas as pd, os, numpy as np, main_pandas
from datetime import datetime as dt

arrPath = main_pandas.arrPath
strSALinput = main_pandas.strSALinput
strSALOutput = main_pandas.strSALoutput
strSUFFIX = main_pandas.strSUFFIX
arrSALheader = main_pandas.arrSALheader


# load Excel
strFilePath = os.path.join(*arrPath, strSALinput)
dfSIF = pd.read_excel(strFilePath, sheet_name='SIF')
dfINT = pd.read_excel(strFilePath, sheet_name='INT')
dfPFE = pd.read_excel(strFilePath, sheet_name='PFE')
dfSIF = dfSIF.drop(arrSALheader, axis=0)

arrSIFcols = ['name_SIF','ref_SIF','desc_SIF','tarRRF_SIF', 'tarSIL_SIF']
arrINTcols = ['C1_INT', 'C2_INT', 'C3_INT', 'C4_INT', 'C5_INT'] + ['vot_INT']
arrPFEcols = ['C1_PFE', 'C2_PFE', 'C3_PFE', 'C4_PFE', 'C5_PFE', 'C6_PFE', 'C7_PFE'] + ['vot_PFE']
arrInstcols = arrINTcols[:-1]+arrPFEcols[:-1]

dfSIF_Rep = dfSIF.reindex(columns=arrSIFcols)
dfSIF_Rep['tarSIL_SIF'] = dfSIF_Rep['tarSIL_SIF'].str[-1]
dfSIF_Rep['Index'] = dfSIF_Rep.index+1
dfSubsys = dfSIF.reindex(columns=arrSIFcols +  arrINTcols + arrPFEcols)


# Remove Asterisk from INT and PFE
dfINT['INT'] = dfINT['INT'].str.replace('*','',regex=False)
dfPFE['PFE'] = dfPFE['PFE'].str.replace('*','',regex=False)


# Lookup Instrument Model
def Replace_Col_with_Asterisk(dfReplace, dfLookup, strCol_Lkp, strCol_Rep, numMaxAsterisk=5):
    dfOutput = dfReplace.copy()
    for intSeq in range (numMaxAsterisk,-1,-1):
        srsHead = pd.Series(['*' * intSeq] * dfLookup[strCol_Lkp].shape[0], index=dfLookup[strCol_Lkp].index)
        arrCol_Lkp = srsHead.str.cat(dfLookup[strCol_Lkp],join='left').to_list()
        arrCol_Rep = srsHead.str.cat(dfLookup[strCol_Rep],join='left').to_list()
        dfOutput = dfOutput.replace(arrCol_Lkp, arrCol_Rep)
    return dfOutput


dfSubsys_Rep = Replace_Col_with_Asterisk(dfSubsys, dfINT, 'INT', 'Model', 5)
dfSubsys_Rep = Replace_Col_with_Asterisk(dfSubsys_Rep, dfPFE,'PFE', 'Model', 5)


def Parse_Instrument_Group_Name (srsInput, isList=False):
    # Get Voting and Instrument as separate List
    arrVote = srsInput[-1].replace(' ','').split(',')
    intVote_Len = len(arrVote)
    arrInst = srsInput[:-1].dropna().to_list()
    # Sort the Instrument and Voting Lists, and put into a dict, based on number of asterisk
    dictMainGrp = dict()
    for intSeq in range(0,intVote_Len):
        setSubGrp = set()
        dictMainGrp[intSeq] = setSubGrp
        for elemInst in arrInst:
            if elemInst.count('*') == intSeq:
                setSubGrp.add(elemInst.replace('*',''))
    # Concat as string for the whole subsystem
    for k, v in dictMainGrp.items():
        strInst = '(' + '&'.join(list(v)) + ')' if (len(list(v)) != 0) else ''
        strVote = arrVote[k-1].replace('*','')
        strInstVote = f'{strVote}{strInst}'
        dictMainGrp[k] = strInstVote
    if isList:
        return list(dictMainGrp.values())
    else:
        return '*'.join(dictMainGrp.values())


dfSIF_Rep['INT'] = dfSubsys_Rep.reindex(columns=arrINTcols).apply(lambda x: Parse_Instrument_Group_Name(x,isList=False), axis=1)
dfSIF_Rep['PFE'] = dfSubsys_Rep.reindex(columns=arrPFEcols).apply(lambda x: Parse_Instrument_Group_Name(x,isList=False), axis=1)
dfSIF_Rep['Changed'] = (dfSubsys.loc[:,arrInstcols] == dfSubsys_Rep.loc[:,arrInstcols]).apply(lambda x: ~x.any(), axis=1)

dfINT_Rep = pd.concat([dfSIF['name_SIF'],dfSIF_Rep['Changed'],dfSIF_Rep['INT'].str.split('*',expand=True)],axis=1)
dfINT_Rep = dfINT_Rep.set_index(['name_SIF','Changed']).stack(0).reset_index().drop(columns='level_2')

dfPFE_Rep = pd.concat([dfSIF['name_SIF'],dfSIF_Rep['Changed'],dfSIF_Rep['PFE'].str.split('*',expand=True)],axis=1)
dfPFE_Rep = dfPFE_Rep.set_index(['name_SIF','Changed']).stack(0).reset_index().drop(columns='level_2')


def Get_Index_For_SIF(dfInput):
    srsShift = dfInput['name_SIF'].shift(1,fill_value='one')
    srsShift = pd.Series((srsShift != dfInput['name_SIF'])*1)
    srsOutput = srsShift.cumsum()
    return srsOutput


dfINT_Rep['Index'] = Get_Index_For_SIF(dfINT_Rep)
dfPFE_Rep['Index'] = Get_Index_For_SIF(dfPFE_Rep)

dfINT_UiPath = dfINT_Rep.copy()
dfINT_UiPath[0] = dfINT_UiPath[0].str.cat(strSUFFIX*dfINT_UiPath.shape[0])
dfINT_UiPath = dfINT_UiPath[(~dfINT_UiPath[0].str.contains('Overall',regex=False)) 
                            & (dfINT_UiPath[0].str.contains('(',regex=False)) 
                            & (dfINT_UiPath[0].str.contains(')',regex=False))]

dfPFE_UiPath = dfPFE_Rep.copy()
dfPFE_UiPath[0] = dfPFE_UiPath[0].str.cat(strSUFFIX*dfPFE_UiPath.shape[0])
dfPFE_UiPath = dfPFE_UiPath[(~dfPFE_UiPath[0].str.contains('Overall',regex=False)) 
                            & (dfPFE_UiPath[0].str.contains('(',regex=False)) 
                            & (dfPFE_UiPath[0].str.contains(')',regex=False))]


# Rename redundant groups in SIFs
def Rename_Redundant_Group(dfInput):
    dfOutput = dfInput.copy()
    srsCount = dfInput.groupby('name_SIF')[0].value_counts()
    srsCount = srsCount[srsCount>1]
    for strSIF, strGroup in srsCount.index.to_list():
        arrIndex = dfOutput[(dfOutput['name_SIF']==strSIF) & (dfOutput[0]==strGroup)].index
        for intSeq, intIndex  in enumerate(arrIndex):
            if intSeq == 0:
                continue
            else:
                dfOutput.at[intIndex,0] = strGroup + f' [{intSeq}]'
    return dfOutput


dfINT_UiPath_Rename = Rename_Redundant_Group(dfINT_UiPath)
dfPFE_UiPath_Rename = Rename_Redundant_Group(dfPFE_UiPath)


dfGrps = pd.concat([dfINT_UiPath_Rename[0].value_counts().sort_index().reset_index(),
                    dfPFE_UiPath_Rename[0].value_counts().sort_index().reset_index()],axis=1)
dfGrps = dfGrps.rename(columns={'index': 'Group',0:'Appearances'})


# Get Redundant Summary
srsINT_redun = dfINT_Rep['Index'].value_counts()
arrINT_redun = srsINT_redun[srsINT_redun>1].index.to_list()
dfINT_Rep = dfINT_Rep[dfINT_Rep['Index'].isin(arrINT_redun)]


srsPFE_redun = dfPFE_Rep['Index'].value_counts()
arrPFE_redun = srsPFE_redun[srsPFE_redun>1].index.to_list()
dfPFE_Rep = dfPFE_Rep[dfPFE_Rep['Index'].isin(arrPFE_redun)]

# Export
strTimeNow = dt.now().strftime("%d%b%y")
with pd.ExcelWriter(os.path.join(*arrPath, f'{strSALOutput}-{strTimeNow}.xlsx')) as writer:
    dfSIF_Rep.to_excel(writer,'Overview')
    dfGrps.to_excel(writer,'GrpSummary')
    dfINT_UiPath_Rename.to_excel(writer,'INT')
    dfPFE_UiPath_Rename.to_excel(writer,'PFE')
    dfINT_Rep.to_excel(writer,'INT_Redun')
    dfPFE_Rep.to_excel(writer,'PFE_Redun')




