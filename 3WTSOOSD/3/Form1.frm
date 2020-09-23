VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMOpen 
      Caption         =   "Multiselect open"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================='
'=                                                            ='
'= ======AUTHOR======                                         ='
'= THIS IS A FREE CODE                                        ='
'= BY FILIP WIELEWSKI                                         ='
'= E-MAIL: WIELFILIST@WP.PL                                   ='
'=                                                            ='
'= ======SORRY FOR:======                                     ='
'= my bad english which I use in descriptions :]              ='
'=                                                            ='
'=============================================================='


'======API Functions======
'Shows "open file" dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
        "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Shows "save file" dialog
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
        "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'======Types======
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000                         'new look commdlg


Private Sub cmdMOpen_Click()
    
    Dim ofnOpen As OPENFILENAME 'description of way to show "open file" dialog
    Dim lonRet As Long          'return value
    Dim strFiles As String      'will store message text
    Dim strArrFiles() As String 'will store path and files' names
    Dim intI As Integer         'variable for "For...Next"
    Dim booMulti As Boolean     'True: there where chosen more than one files;...
                                '...False: only one file was chosen
                                
    With ofnOpen
        'set size
        .lStructSize = Len(ofnOpen)
        'set "mother" object
        .hwndOwner = Me.hWnd
        'set instance of application (our application)
        .hInstance = App.hInstance
        'set filter
        .lpstrFilter = "Text files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + _
        "VB project files (*.vbp)" + Chr$(0) + "*.vbp" + Chr$(0) + _
        "All files (*.*)" + Chr$(0) + "*.*"
        'set buffer
        .lpstrFile = String(1024, 0)
        'set maximum number of chars
        .nMaxFile = 1024
        'set buffer
        .lpstrFileTitle = String(1024, 0)
        'set maximum number of chars
        .nMaxFileTitle = 1024
        'set initial directory
        .lpstrInitialDir = "C:\Program files"
        'set title (it will be displayed on "open dialog"
        .lpstrTitle = "Here is the title!!!"
        'extra flags: allow multiselect and "new appearance" (see how will...
        '..."open dialog" look like without OFN_EXPLORER parameter - like in...
        '...old windows...)
        .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    End With
    
    'IMPORTANT: when user choose a lot of files and their names will be to...
    '...long then GetOpenFileName function will fail. This is the...
    '...reason of that our buffer is set to 1024 chars (not to 255).
    
    'if you choose more than one file then lpstrFile argument stores:...
    '...<path> and Chr$(0) and <file1 name> and Chr$(0) and <file2 name>...
    '... and Chr$(0) & <file3 name>... etc.
    
    'show "open file" dialog
    lonRet = GetOpenFileName(ofnOpen)
    
    'If function succeeded then:
    If lonRet Then
    
        'check if there was chosen only one file or there were chosen...
        '...more files
        If InStr(1, ofnOpen.lpstrFile, Chr$(0), vbTextCompare) = 4 Then
            booMulti = True
        ElseIf Mid$(ofnOpen.lpstrFile, InStr(1, ofnOpen.lpstrFile, Chr$(0), _
        vbTextCompare) + 1, 1) <> Chr$(0) Then
            booMulti = True
        Else
            booMulti = False
        End If
        
        'if there were choosen more than one files then:
        If booMulti = True Then
            'put directory and names of files into strArrFiles array
            strArrFiles = Split(ofnOpen.lpstrFile, Chr$(0))
            'set message's text
            strFiles = "Directory of files you chose is:" & vbCrLf & "  " & Chr$(34) & _
            strArrFiles(0) & Chr$(34) & vbCrLf & "Files:"
            'get name of all files and (if their names <> "") then put them into...
            '...strFiles
            For intI = 1 To UBound(strArrFiles)
                If strArrFiles(intI) <> "" Then strFiles = strFiles & vbCrLf & _
                "  " & Chr$(34) & strArrFiles(intI) & Chr$(34)
            Next intI
            'display message with selected files
            MsgBox strFiles, vbInformation
        'if there was chosen only one file then:
        Else
            'Display message with selected file
            MsgBox "You sellected file " & Chr$(34) & Left$(ofnOpen.lpstrFile, _
            InStr(1, ofnOpen.lpstrFile, Chr$(0)) - 1) & Chr$(34), vbInformation
        End If
        
    'User probably pressed "cancel" (or error occurred)
    Else
        MsgBox "You pressed " & Chr$(34) & "Cancel" & Chr$(34), vbInformation
    End If
    
End Sub

Private Sub cmdOpen_Click()
    
    Dim ofnOpen As OPENFILENAME 'description of way to show "open file" dialog
    Dim lonRet As Long          'return value

    With ofnOpen
        'set size
        .lStructSize = Len(ofnOpen)
        'set "mother" object
        .hwndOwner = Me.hWnd
        'set instance of application (our application)
        .hInstance = App.hInstance
        'set filter
        .lpstrFilter = "Text files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + _
        "VB project files (*.vbp)" + Chr$(0) + "*.vbp" + Chr$(0) + _
        "All files (*.*)" + Chr$(0) + "*.*"
        'set buffer
        .lpstrFile = String(255, 0)
        'set maximum number of chars
        .nMaxFile = 255
        'set buffer
        .lpstrFileTitle = String(255, 0)
        'set maximum number of chars
        .nMaxFileTitle = 255
        'set initial directory
        .lpstrInitialDir = "C:\Program files"
        'set title (it will be displayed on "open dialog"
        .lpstrTitle = "Here is the title!!!"
        'no extra flags
        .flags = 0
    End With
    
    'IMPORTANT: when path of file chosen by user has more than 255 letters...
    '...then GetOpenFileName fails. To avoid this situation...
    '...create longer buffer.
    
    'show "open file" dialog
    lonRet = GetOpenFileName(ofnOpen)
    'If function succeeded then:
    If lonRet Then
        'Display message with selected file
        MsgBox "You sellected file " & Chr$(34) & Left$(ofnOpen.lpstrFile, _
        InStr(1, ofnOpen.lpstrFile, Chr$(0)) - 1) & Chr$(34), vbInformation
    Else
        'User probably pressed "cancel"
        MsgBox "You pressed " & Chr$(34) & "Cancel" & Chr$(34), vbInformation
    End If
    
End Sub

Private Sub cmdSave_Click()
    
    Dim ofnSave As OPENFILENAME 'description of way to show "save file" dialog
    Dim lonRet As Long          'return value

    With ofnSave
        'set size
        .lStructSize = Len(ofnSave)
        'set "mother" object
        .hwndOwner = Me.hWnd
        'set instance of application (our application)
        .hInstance = App.hInstance
        'set filter
        .lpstrFilter = "VB project files (*.vbp)" + Chr$(0) + "*.vbp"
        'set buffer
        .lpstrFile = String(255, 0)
        'set maximum number of chars
        .nMaxFile = 255
        'set buffer
        .lpstrFileTitle = String(255, 0)
        'set maximum number of chars
        .nMaxFileTitle = 255
        'set initial directory
        .lpstrInitialDir = "C:\Program files"
        'set title (it will be displayed on "save dialog"
        .lpstrTitle = "Here is the title!!!"
        'no extra flags
        .flags = 0
    End With
    
    'IMPORTANT: when path of file chosen by user has more than 255 letters...
    '...then GetSaveFileName fails. To avoid this situation...
    '...create longer buffer.
    
    'show "save file" dialog
    lonRet = GetSaveFileName(ofnSave)
    'If function succeeded then:
    If lonRet Then
        'Display message with selected file
        MsgBox "You sellected file " & Chr$(34) & Left$(ofnSave.lpstrFile, _
        InStr(1, ofnSave.lpstrFile, Chr$(0)) - 1) & Chr$(34), vbInformation
    Else
        'User probably pressed "cancel"
        MsgBox "You pressed " & Chr$(34) & "Cancel" & Chr$(34), vbInformation
    End If
    
End Sub
