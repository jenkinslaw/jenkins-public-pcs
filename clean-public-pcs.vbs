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
arrFolders = Array("B:\")


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

Sub DeleteSubfolders(strPath)
    'Deletes all subfolders in a given folder, strPath.
    '
    'strPath takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder\
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\

    Dim objFSO, objFolder, arrSubfolders, f, strSubfolder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getFolder(strPath)
    Set arrSubfolders = objFolder.subFolders
    
    For Each f In arrSubfolders
        'Replace the first instance of strPath in f with ""
        'This will result in just the subfolder name.
        strSubfolder = Replace(f,strPath,"",1,1)
        
        'Do not delete system folders, like "$RECYCLE.BIN"
        'Some users don't have permission to even attempt this, so scipt fails.
        If left(strSubfolder, 1) <> "$" And _
        strSubfolder <> "System Volume Information" Then
            objFSO.DeleteFolder(f)
        End If
    Next
    
End Sub

'---------------------------------------------------------------
'DELETE FILES
'---------------------------------------------------------------
Dim i

For Each i In arrFolders
    
    'Ensure that the filepath ends with "\*.*"
    If right(i, 1) <> "\" Then
        i = i & "\"
    End If
    
    DeleteSubfolders i 'Delete subfolders in arrFolders[i]
    DeleteFiles i & "*.*" 'Delete all types of files in arrFolders[i]
    
    
Next