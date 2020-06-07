
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

# Relative path of the input file
arrPath = ['..','02-200601-Project']

# General Unit Name
strUnit = 'EPS'

# File name of user input file
strSALinput = strUnit + '-Import-05Jun20' + '.xlsx'
strSADinput = strUnit + '-SIL Verification 6-5-2020 SILver Detailed Report' + '.docx'
strSAMinput = strUnit + '-SADetail-DataBase' + '.sdb'
strSACinput = strSADinput

#-----------------------------------------------------------------------------------------------------------------------
# SAList Specific Input

# SAList file Details
strSALtab = 'SIF'
strTagLookup = 'name_SIF'

# If there is one row of header for MailMerge then put 0, if not put 'None'
# excluding the row necessary to make MailMerge works
intSALheader = 0

# User parameter for PTI
strSUFFIX = ['_3Y']

#-----------------------------------------------------------------------------------------------------------------------
# Name of Output Identifier (Change only if you know what you doing)
strSALoutput = strUnit + '-SAList'
strSADoutput = strUnit + '-SADetail'
strSAMoutput = strUnit + '-SAMerge'
strSACoutput = strUnit + '-SAComps'

# File name of General Param Inputs (Change only if you know what you doing)
strSASparam = 'SAS-Params-12Nov19' + '.xlsx'
strSAMtmple = 'SAM_Tmplt-12May20' + '.xlsm'
isDebug = False

#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

import sys, os
if __name__ == '__main__':
    sys.path.append(os.getcwd())
    arrPathShort = os.getcwd().split(os.path.sep)[-2:]
    intUI = int(input('Which script you would like to run? [SAList=1, SADetail=2, SAMerge=3, SAComps=4]: '))
    if intUI == 1:
        strPath_Script = os.path.join(*arrPathShort, 'SAList.py')
        print(f'{strPath_Script} is being ran on {os.path.join(*arrPath, strSALinput)}.')
        import SAList
        SAList.__name__
    elif intUI == 2:
        strPath_Script = os.path.join(*arrPathShort, 'SADetail.py')
        print(f'{strPath_Script} is being ran on {os.path.join(*arrPath, strSADinput)}.')
        import SADetail
        SADetail.__name__
    elif intUI == 3:
        strPath_Script = os.path.join(*arrPathShort, 'SAMerge.py')
        print(f'{strPath_Script} is being ran on {os.path.join(*arrPath, strSAMinput)}.')
        import SAMerge
        SAMerge.__name__
    elif intUI == 4:
        strPath_Script = os.path.join(*arrPathShort, 'SAComps.py')
        print(f'{strPath_Script} is being ran on {os.path.join(*arrPath, strSACinput)}.')
        import SAComps
        SAComps.__name__
    else:
        print('Invalid Input: Script Exit.')
        input('Press any key to exit.')
        exit()

    print(f'{strPath_Script} has ran successfully.')
    input('Press any key to exit.')

