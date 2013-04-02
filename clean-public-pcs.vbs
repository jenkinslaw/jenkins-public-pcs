'---------------------------------------------------------------
'CLEAN UP PUBLIC PC
'---------------------------------------------------------------
'Deletes all files from the user's My Documents folder. Also deletes
'everything on the user's desktop.
'
'Run as a logout script for SAM.

Option Explicit

'---------------------------------------------------------------
'CONFIGURATION OPTIONS
'---------------------------------------------------------------
'Delcaring / setting configurable variables.
Dim arrFolders

'Each element of arrFolders represents the path to a folder that will be
'emptied. To target specific files, use wildcards like *.txt.
arrFolders = Array("B:\*.*", "C:\Users\stn-5\Desktop\*.*")


'Should point to B drive in reality

'---------------------------------------------------------------
'SUBS
'---------------------------------------------------------------
Sub DeleteFiles(strPath)
    'Useful for deleting more than one file that match a wildcard criteria
    'within strPath.
    '
    'Empties contents of strPath.
    'Path can contain a full filename or wildcards, like "*.*" or "*.txt"
    '
    'strPath takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\*.txt
    '   D:\Folder\Subfolder\
    '   \\network-path\subfolder\*.*
    '   \\network-path\c$\subfolder\
    
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    objFSO.DeleteFile(strPath)
End Sub

Sub DeleteFolders(strPath)
    'Deletes all subfolders in a given folder, strPath.
    '
    'Path can contain a full filename or wildcards, like "*.*" or "*.txt"
    '
    'strPath takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder\
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\

    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    objFSO.DeleteFolder(strPath)
    
End Sub

'---------------------------------------------------------------
'DELETE FILES
'---------------------------------------------------------------
Dim i

For Each i In arrFolders
    
    'Ensure that the filepath ends with "\*.*"
    If right(i, 1) <> "\" Then
        i = i & "\*.*"
    Else
        i = i * "*.*"
    End If
    
    DeleteFiles i
    DeleteFolders i
    
Next