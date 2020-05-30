import os


def strSpecificFileName(strMyFile, strExtension, strPath='./', isRequired=True):
    """Return filename (string) of one specific file based on 'Generic Name' and its extension.
    The script exit when multiple file or no file found, if the file is flagged as necessary."""
    seqFile = 0
    arrFileDir = os.listdir(strPath)
    for strFile in arrFileDir:
        if strFile.find(strMyFile) != -1 and strFile.find(strExtension) != -1:
            seqFile += 1
            temp = strFile
    if seqFile == 1:
        print(f"File '{strPath + temp}' Found.")
        return strPath + temp
    elif isRequired:
        if seqFile == 0:
            print(f"Error: No File '{strMyFile}' with Extension '.{strExtension}' Found.")
        elif seqFile > 1:
            print(f"Error: Multiple Files '{strMyFile}' with Extension '.{strExtension}' Found "
                  f"or The File is Opened.")
        _ = input("Press Any Key to Exit.")
        exit()
    else:
        print(f"Specific File '{strMyFile}' with Extension '.{strExtension}' Not Found, Operation Passed.")
        return None


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
