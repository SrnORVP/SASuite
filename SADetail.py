
import os, docx, pandas as pd, numpy as np, pickle as pk, main_pandas
from datetime import datetime as dt
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

arrPath = main_pandas.arrPath
strSADinput = main_pandas.strSADinput
arrPath = main_pandas.arrPath
strSASparam = main_pandas.strSASparam
strSADoutput = main_pandas.strSADoutput
isDebug = main_pandas.isDebug

# Docx Blocks Generator
def iter_Docx_Block_Items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent. _tc
    else:
        raise ValueError("Something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# Extract Word Docx
def extract_Block_From_Docs(docxSIF):
    arrBlkSeq = [] 
    arrBlkStyle = []
    arrObjBlk = []
    for seqBlk, objBlk in enumerate(iter_Docx_Block_Items(docxSIF)):
        if isinstance(objBlk, Paragraph):
            if objBlk.text and not objBlk.text.isspace():
                arrBlkSeq.append(seqBlk)
                arrBlkStyle.append(objBlk.style.name)
                arrObjBlk.append(objBlk.text.strip().replace("  ", " "))
        else:
            for objBlkCell in objBlk._cells:
                arrObjBlk.append(objBlkCell.text)
                arrBlkSeq.append(seqBlk)
                arrBlkStyle.append('Cell')
    return arrBlkSeq, arrBlkStyle, arrObjBlk


# Remove Head and Tail of Blocks Extracted
def remove_Array_HeadTail(arrBlkSeq, arrBlkStyle, arrObjBlk):
    idxBlkArrHead = arrObjBlk.index('Project Description:')+2
    for intSeq, elem in enumerate(arrObjBlk):
        if "Group Reuse Overview" in elem:
            idxBlkArrTail = intSeq
    arrBlkSeq = arrBlkSeq[idxBlkArrHead:idxBlkArrTail]
    arrBlkStyle = arrBlkStyle[idxBlkArrHead:idxBlkArrTail]
    arrObjBlk = arrObjBlk[idxBlkArrHead:idxBlkArrTail]
    return arrBlkSeq, arrBlkStyle, arrObjBlk


# String Operation for 'Left', 'Right' and 'Mid' in Excel
def extract_Substring(strMainString, strAction, strSubstring1, strSubstring2=None):
    if strSubstring1 in strMainString:
        idxSubStr1 = strMainString.index(strSubstring1)
    else:
        return strMainString
    lenSubstring = len(strSubstring1)
    idxStart = 0
    idxEnd = None
    if strAction.lower() == 'left_string':
        idxEnd = idxSubStr1
    elif strAction.lower() == 'right_string':
        idxStart = idxSubStr1 + lenSubstring
    elif strAction.lower() == 'mid_string' and strSubstring2 != None:
        idxSubStr2= strMainString.index(strSubstring2)
        idxStart = idxSubStr1 + lenSubstring
        idxEnd = idxSubStr2
    return strMainString[idxStart:idxEnd]


# Create a Column for Identifier Based on Style in Docx
def assign_Identifier_By_Docx_Style(dfInput, arrID, strLkpCol, strDataCol):
    for seqRowID, rowID in enumerate(arrID):
        dfInput[rowID[0]] = dfInput[strDataCol].apply(extract_Substring, strAction = rowID[1], strSubstring1 = rowID[2]).where(dfInput[strLkpCol]==rowID[0])
        if seqRowID != 0:
            dfInput[rowID[0]] = dfInput[rowID[0]].fillna(dfInput[arrID[seqRowID-1][0]])
    assign_Identifier_By_Block_Content(dfInput, 'Logic Solver Name:')    
    arrIDidxCol = arrID[:,0]
    dfInput.loc[:,arrIDidxCol] = dfSAD.loc[:,arrIDidxCol].ffill()
    return dfInput


# Additional Identifier for LS
def assign_Identifier_By_Block_Content(dfInput, strFind):
    dfInput['Heading 5'] = dfInput['blk'].shift(-1).where(dfInput['blk']==strFind)
    dfInput['Heading 4'] = dfInput['Heading 4'].fillna(dfInput['Heading 5'])
    return dfInput


# Create a Column Counter on Row for DataFrame
def assign_Col_num_for_DF_entry(dfInput, strLkpCol, strColName, isInplace=False):
    ndaLkpCol = dfInput[strLkpCol].tolist()
    arrOutput = [1]
    seqDF_Cnt = 1
    for seqDF in range(1,len(dfInput)):
        if ndaLkpCol[seqDF] == ndaLkpCol[seqDF-1]:
            seqDF_Cnt += 1
        else:
            seqDF_Cnt = 1
        arrOutput.append(seqDF_Cnt)
    if isInplace:
        dfInput[strColName] = arrOutput
        return dfInput
    else:
        dfOutput = dfInput.copy()
        dfOutput[strColName] = arrOutput
        return dfOutput


# Create a Index Counter on Rows in DataFrame
def assign_idx_for_DF_entry(dfInput, strLkpCol1, strLkpCol2, strColName, isInplace=False):
    ndaLkpCol1 = dfInput[strLkpCol1].tolist()
    ndaLkpCol2 = dfInput[strLkpCol2].tolist()
    arrOutput = [1]
    seqDF_Cnt = 1
    for seqDF in range(1,len(dfInput)):
        if (ndaLkpCol1[seqDF] != ndaLkpCol1[seqDF-1]) or (ndaLkpCol2[seqDF] != ndaLkpCol2[seqDF-1]):
            seqDF_Cnt += 1
        arrOutput.append(seqDF_Cnt)
    if isInplace:
        dfInput[strColName] = arrOutput
        return dfInput
    else:
        dfOutput = dfInput.copy()
        dfOutput[strColName] = arrOutput
        return dfOutput


# Extract Value based on criteria
def extract_Value_Col_on_Criteria(dfInput, arrCriteria, intShift=1):
    dfFinal = pd.DataFrame()
    for rowCriteria in arrCriteria:
        dfOutput = dfInput[(dfInput==rowCriteria[0]).shift(intShift, axis=1, fill_value=False)]
        arrExt = []
        for seqRow in range(0,len(dfOutput)):
            srsOutput = dfOutput.iloc[seqRow]
            srsOutput = srsOutput.dropna()
            if srsOutput.empty:
                arrExt.append('-')
            else:
                arrExt.append(srsOutput.iloc[0])
        dfFinal[rowCriteria[1]] = arrExt
    return dfFinal


def extract_subsys_result(dfInput, arrSubsys, arrCriteria):
    # extract all entry with the same identifier
    dfInput_Subsys = dfInput[(dfSAD['Heading 3']!=dfSAD['Heading 4']) & (dfSAD['Heading 3']==arrSubsys[0])]
    # indexing entry for pivoting
    dfInput_Subsys = assign_idx_for_DF_entry(dfInput_Subsys,'Heading 1', 'Heading 4', arrSubsys[0])
    dfInput_Subsys = assign_Col_num_for_DF_entry(dfInput_Subsys,arrSubsys[0], arrSubsys[1])
    # Pivot each entry as column
    dfInput_Subsys = dfInput_Subsys.pivot(index=arrSubsys[0], values='blk', columns=arrSubsys[1])
    dfInput_Subsys = dfInput_Subsys.drop_duplicates()
    # reindexing
    dfInput_Subsys[1] = dfInput_Subsys[1].apply(extract_Substring, strAction='right_string', strSubstring1=': ')
    dfInput_Subsys = dfInput_Subsys.set_index(dfInput_Subsys[arrSubsys[3]])
    # extract those columns that have the desired result
    dfOutput_Subsys = extract_Value_Col_on_Criteria(dfInput_Subsys,arrCriteria)
    dfOutput_Subsys.index = dfInput_Subsys.index
    return dfOutput_Subsys


# Extract SIF tag corresponding to Device
def extract_SIF_Device_Pair(dfInput, strCond, strColname):
    dfOutput = dfInput[(dfInput['Heading 3']!=dfInput['Heading 4']) & (dfInput['Heading 3']==strCond)]
    dfOutput = dfOutput.loc[:,['Heading 1', 'Heading 4']].drop_duplicates()
    dfOutputPivot = dfOutput.sort_values(by=['Heading 4','Heading 1'], axis=0)
    dfOutputPivot = assign_Col_num_for_DF_entry(dfOutputPivot,'Heading 4',strColname)
    dfOutputPivot = dfOutputPivot.pivot(index='Heading 4', columns=strColname, values='Heading 1')
    dfOutput.index.name = ''
    dfOutputPivot.index.name = ''
    return dfOutput, dfOutputPivot


# Join SIF-Device pair with Device Attribute and return Attribute DataFrame for every SIF
def extract_PTI_from_subsys(dfInput, strLkpCol, dfAttr, strAttr):
    dfAttr = dfAttr[strAttr]
    dfOutput = dfInput.join(dfAttr,on=strLkpCol,how='left')
    dfOutput = dfOutput.drop(labels=strLkpCol,axis=1).drop_duplicates()
    dfOutput = assign_Col_num_for_DF_entry(dfOutput,'Heading 1', 'cntPTIcol')
    dfOutput = dfOutput.pivot(index='Heading 1', values=strAttr, columns='cntPTIcol')
    dfOutput = dfOutput.apply(concat_DF_Rows,axis=1,strConcat=', ')
    return dfOutput


# Concat DataFrame Rows
def concat_DF_Rows(srsInput, strConcat):
    arrOutput = srsInput.to_numpy()
    strOutput = arrOutput[0]
    for seqOutput in range(0,len(arrOutput)):
        if isinstance(arrOutput[seqOutput],str) and seqOutput != 0:
            strOutput = strOutput + ', ' + arrOutput[seqOutput]
        else:
            strOutput
    return strOutput


# Load Inputs
docxSIF = docx.Document(os.path.join(*arrPath, strSADinput))
arrC4SAD_lvl = pd.read_excel(os.path.join('.', strSASparam), 'C4SAD-lvl').values
arrC4SAD_col = pd.read_excel(os.path.join('.', strSASparam), 'C4SAD-col').values
arrC4SAD_grp = pd.read_excel(os.path.join('.', strSASparam), 'C4SAD-grp').values
arrC4SAD_sub = pd.read_excel(os.path.join('.', strSASparam), 'C4SAD-sub').values

arrBlkSeq, arrBlkStyle, arrObjBlk = extract_Block_From_Docs(docxSIF)
arrBlkSeq, arrBlkStyle, arrObjBlk = remove_Array_HeadTail(arrBlkSeq, arrBlkStyle, arrObjBlk)
dfSAD = pd.DataFrame({'style': arrBlkStyle, 'blk':arrObjBlk }, index = arrBlkSeq)
dfSAD = dfSAD.drop(index = dfSAD[dfSAD['style']=='Caption'].index)
dfSAD = assign_Identifier_By_Docx_Style(dfSAD, arrC4SAD_lvl, 'style', 'blk')


# Overall SIF Result
dfSAD_SIF = dfSAD[(dfSAD['Heading 3']==dfSAD['Heading 4'])]
dfSAD_SIF = assign_Col_num_for_DF_entry(dfSAD_SIF,'Heading 1', 'count')

dfSAD_SIF = dfSAD_SIF.pivot(index='Heading 1', values='blk', columns='count')

arrSAD_SIF_ftdCol = arrC4SAD_col[:,1]
arrSAD_SIF_dict = dict(arrC4SAD_col[:,1:])

dfSAD_SIF = dfSAD_SIF.loc[:,arrSAD_SIF_ftdCol].rename(columns=arrSAD_SIF_dict)


# Individual Subsystem Result
# Extract desire result from each subsystem as DataFrame
dictSADgrp = {}
for rowC4SADsub in arrC4SAD_sub:
    arrGrp_Crit = arrC4SAD_grp[arrC4SAD_grp[:,0]==rowC4SADsub[0]][:,1:]
    dfIndGrp = extract_subsys_result(dfSAD, rowC4SADsub, arrGrp_Crit)
    dfIndGrp = dfIndGrp.sort_index(axis=0)
    dictSADgrp[rowC4SADsub[0]] = dfIndGrp

dfSAD_INTgrp = dictSADgrp[arrC4SAD_sub[0][0]]
dfSAD_PFEgrp = dictSADgrp[arrC4SAD_sub[1][0]]
dfSAD_PLCgrp = dictSADgrp[arrC4SAD_sub[2][0]]


# Get a List of SIF with corresponding Sensor or Final Element Group
dfSAD_INTpair, dfSAD_INTkeys = extract_SIF_Device_Pair(dfSAD, 'Sensor', 'cntINTkeys')
dfSAD_PFEpair, dfSAD_PFEkeys = extract_SIF_Device_Pair(dfSAD, 'Final Element', 'cntPFEkeys')


# Get PTI for Initator and Final Element
dfSAD_INTpair = extract_PTI_from_subsys(dfSAD_INTpair, 'Heading 4', dfSAD_INTgrp, 'PTI_INTgrp')
dfSAD_PFEpair = extract_PTI_from_subsys(dfSAD_PFEpair, 'Heading 4', dfSAD_PFEgrp, 'PTI_PFEgrp')

dfSAD_SIF['PTI_INT'] = dfSAD_INTpair
dfSAD_SIF['PTI_PFE'] = dfSAD_PFEpair


# Exporting Result
strTimeNow = dt.now().strftime("%d%b%y")
with pd.ExcelWriter(f'{os.path.join(*arrPath, "")}{strSADoutput}-{strTimeNow}.xlsx') as writer:
    dfSAD_SIF.to_excel(writer,'SIFs')
    dfSAD_INTgrp.to_excel(writer,'INT Device')
    dfSAD_INTkeys.to_excel(writer,'INT Keys')
    dfSAD_PFEgrp.to_excel(writer,'PFE Device')
    dfSAD_PFEkeys.to_excel(writer,'PFE Keys')
    dfSAD_PLCgrp.to_excel(writer,'PLC Device')

if isDebug:
    with pd.ExcelWriter(f'{strSADoutput}-debug-{strTimeNow}.xlsx') as writer:
        dfSAD.to_excel(writer, 'SAD')


# Generate SADetail Pickle File for SAMerge
strPkName = f'{os.path.join(*arrPath, "")}{strSADoutput}-DataBase.sdb'
with open(strPkName, 'xb') as pkWriter:
    pk.dump(dfSAD_SIF, pkWriter)

