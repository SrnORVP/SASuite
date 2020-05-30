
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

# Relative path of the input file
arrPath = ['..','01-Test']

# General Unit Name
strUnit = 'AUX PDH'

# File name of user input file
strSALinput = strUnit + '-Import-28May20' + '.xlsx'
strSADinput = strUnit + '-SIL Verification Study 5-27-2020 SILver Detailed Report' + '.docx'
strSAMinput = strUnit + '-SADetail-DataBase' + '.sdb'
strSACinput = strSADinput

# Name of Output Identifier
strSALoutput = strUnit + '-SAList'
strSADoutput = strUnit + '-SADetail'
strSAMoutput = strUnit + '-SAMerge'
strSACoutput = strUnit + '-SAComps'

# File name of General Param Inputs
strParaFile = 'SASuiteInput-12Nov19' + '.xlsx'
strSAMtp = 'SAM_Template-12May20' + '.xlsm'
isDebug = False

#-----------------------------------------------------------------------------------------------------------------------
# SAList Specific Input

# SAList file Details
strSALtab = 'SIF'
strTagLookup = 'name_SIF'
arrSALheader = 0

# User parameter for PTI
strSUFFIX = ['_3Y']

#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

import sys, os
if __name__ == '__main__':
    sys.path.append(os.getcwd())
    arrPathShort = os.getcwd().split(os.path.sep)[-2:]
    intUI = int(input('Which script you would like to run? [SAList=1, SADetail=2, SAMerge=3, SAComps=4]: '))
    if intUI==1:
        strPath_Script = os.path.join(*arrPathShort, 'SAList.py')
        print(f'{strPath_Script} is being ran.')
        import SAList
        print(SAList.__name__)
    elif intUI==2:
        strPath_Script = os.path.join(*arrPathShort, 'SADetail.py')
        print(f'{strPath_Script} is being ran.')
        import SADetail
        print(SADetail.__name__)
    elif intUI==3:
        strPath_Script = os.path.join(*arrPathShort, 'SAMerge.py')
        print(f'{strPath_Script} is being ran.')
        import SAMerge
        print(SAMerge.__name__)
    elif intUI==4:
        strPath_Script = os.path.join(*arrPathShort, 'SAComps.py')
        print(f'{strPath_Script} is being ran.')
        import SAComps
        print(SAComps.__name__)
    else:
        print('Invalid Input: Script Exit.')
        input('Press any key to exit.')
        exit()

    print(f'{strPath_Script} has ran successfully.')
    input('Press any key to exit.')

