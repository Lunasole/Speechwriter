Attribute VB_Name = "modF"
'paste the code below into a new module... then use it as listed below
'*****************************************************************************
'**         Orignal Source Code Credit to: Stefan                           **
'**         http://www.freevbcode.com/ShowCode.Asp?ID=822                   **
'**                                                                         **
'**         Very few modifications made by Brent Coppock                    **
'**         The only change is the transportability of the                  **
'**         code.  Its is now in a module and can be called                 **
'**         from anywhere.  Function signature is slightly                  **
'**         different now.                                                  **
'**         Usage:                                                          **
'**         GetFiles pathtosearch, subfolders, listviewobject               **
'**         Where:  PathToSearch is a string                                **
'**                 SearchSubFolders is a Boolean                           **
'**                 ListViewObject is a valid listview object reference     **
'**         TIP:  Before you call the GetFiles function make your           **
'**               listview object invisible then make it visible after the  **
'**               function finishes.  Very nice speed improvement!!         **
'*****************************************************************************

Option Explicit
'Private Const FILE_ATTRIBUTE_READONLY = &H1
'Private Const FILE_ATTRIBUTE_HIDDEN = &H2
'Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Global FileList() As String
Global FileCount As Long
Public Sub GetFiles(Path As String, SubFolder As Boolean)
    Screen.MousePointer = 11
    
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, FName As String
    
    fPath = AddBackslash(Path)
    FName = fPath & "*.*"
    
    'Wir wollen hier mal alle Dateien im angegebenen Verzeichnis suchen.
    hFile = FindFirstFile(FName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        Call AddFileToList(StripNulls(WFD.cFileName))
    End If
    
    While FindNextFile(hFile, WFD)
        'Solange "FindNextFile" ausfuehren, bis keine Datei mehr gefunden wird, also hFile 0 ist.
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            Call AddFileToList(StripNulls(WFD.cFileName))
        End If
    Wend
    
    If SubFolder Then
        
        'Wenn gewuenscht, dann werden hier alle Unterverzeichnisse durchsucht.
        hFile = FindFirstFile(FName, WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
        StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
        
            'Das ist nun der rekursive Aufruf, die Funktion ruft sich selbst mit neuen Argumenten
            '(in diesem Fall das Unterverzeichnis als "Path") auf.
            GetFiles fPath & StripNulls(WFD.cFileName), True
        End If
        
        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
            StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then

                'Wie oben, hier werden alle weiteren Unterverzeichnisse durchsucht
                GetFiles fPath & StripNulls(WFD.cFileName), True
            End If
        Wend
        
    End If
    FindClose hFile

    Screen.MousePointer = 0
End Sub

Private Function StripNulls(f As String) As String
    'Schneidet einen String nach dem ersten "Chr$(0)" ab, und gibt ihn zurueck.
    StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(S As String) As String
    'Fuegt zu einem String einen abschiessenden Backslash hinzu, wenn notwendig.
    If Len(S) Then
       If Right$(S, 1) <> "\" Then
          AddBackslash = S & "\"
       Else
          AddBackslash = S
       End If
    Else
       AddBackslash = "\"
    End If
End Function


Public Function fc_GetNameFromPath(ByVal fPath As String) As String
 fc_GetNameFromPath = Mid$(fPath, InStrRev(fPath, "\", -1&, vbBinaryCompare) + 1&)
End Function

Private Sub AddFileToList(ByRef LstData$)
Const Ext As String = ".txt"
  LstData = LCase$(LstData)
  If Not Right$(LstData, 4&) = Ext Then Exit Sub
  If Not Left$(LstData, 7&) = "source_" Then Exit Sub
  
  FileCount = FileCount + 1&
    If FileCount > UBound(FileList) Then
      ReDim Preserve FileList(1 To (FileCount + 1000&)) As String
    End If
  FileList(FileCount) = LstData
End Sub
