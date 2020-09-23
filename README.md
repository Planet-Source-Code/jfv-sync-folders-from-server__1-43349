<div align="center">

## Sync Folders from Server


</div>

### Description

This will copy folders from one location to another continuously checking if there are new folders created on the server and checking the last date accessed for the server location to copy to the local computer or other location. This is setup for no user input and will run until it is killed by the Task Manager.
 
### More Info
 
Hard code in your locations or you can wrap it in a form to specify the locations needed. This is currently being used to copy numeric folders, but can be modified to use named folders as well.

I hope you like it! Please commment!

It will take up a little memory when running, but should reset after it gets through each overall loop.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JFV](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jfv.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jfv-sync-folders-from-server__1-43349/archive/master.zip)





### Source Code

```
Dim fso
  Dim n As Double 'Use n for the array variable
  Dim TIFolders(100000) As String 'Create an array for the original folders
  Dim AgentFolders(100000) As String 'Create an array for the destination folders
  Dim TIDate 'Designated for the GetFolder command for the original folders
  Dim AgentDate 'Designated for the GetFolder command for the destination folders
  Do While Now() = Now()                         'Set to continuously run until you kill it in Task Manager
    Set fso = CreateObject("Scripting.FileSystemObject") 'Set up to access FileSystemObject properties
    For n = 0 To 99999                           'n can be set to any amount within the arrays
      TIFolders(n) = "\\<SERVER FOLDER LOCATION>\" & n 'Set the location of the original folder
      AgentFolders(n) = "\\<LOCAL OR OTHER SERVER LOCATION>\" 'Set the destination location of the copy
      If fso.FolderExists(TIFolders(n)) Then 'Checking if the folder exists in the original location
        If fso.FolderExists(AgentFolders(n) & n) Then 'Checking to see if the folder exists on the destination location
          Set TIDate = fso.GetFolder(TIFolders(n)) 'Gets the folder information from the original location
          Set AgentDate = fso.GetFolder(AgentFolders(n)) 'Gets the folder information from the destination location
          If TIDate.DateLastAccessed < AgentDate.DateLastAccessed Then 'If the original location was accessed before
                                                              'the destination location,Then nothing...
          Else
            fso.DeleteFolder AgentFolders(n) & n, True 'Delete destination location
            fso.CopyFolder TIFolders(n), AgentFolders(n) & "\" 'Copy original location to the destination location
          End If
        Else
          fso.CopyFolder TIFolders(n), AgentFolders(n) & "\" 'Otherwise, just copy the original location to the destination location
        End If
      End If
    Next n 'Go to the next folder in the array
    Set fso = Nothing 'Destroy to free memory
    Set TIDate = Nothing 'Destroy to free memory
    Set AgentDate = Nothing 'Destroy to free memory
  Loop 'Start the whole process again...
```

