VERSION 5.00
Begin VB.Form cmdGet 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.FileListBox filFile 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.DirListBox dirDir 
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "cmdGet"
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

'This way of getting file's path is very easy and probably known for everyone:...
'...It uses three controls: DirListBox, DriveListBox and FileListBox.
'drvDrive control shows drives on your computer
'dirDir control shows folders on drive selected in drvDrive control
'filFile control shows files in folder selected in dirDir control


Private Sub Command1_Click()
    
    'If no file in filFile is selected then:
    If filFile.FileName = "" Then
        'display message
        MsgBox "No file selected.", vbExclamation
        'and exit sub
        Exit Sub
    End If
    
    'Get path of file selected in filFile
    MsgBox dirDir.Path & "\" & filFile.FileName
    
End Sub

Private Sub dirDir_Change()
    
    'Connect filFile control with dirDir control
    filFile.Path = dirDir.Path
    
End Sub

Private Sub drvDrive_Change()
    
    'some drives may be unavailable so avoid errors and closing our application
    On Error GoTo ErrDrive
    
    'Connect drvDrive control with dirDir control
    dirDir.Path = drvDrive.Drive
    
    'Exit sub
    Exit Sub
    
ErrDrive:
    'Display message with error description
    MsgBox "Error number: " & Err.Number & vbCrLf & "Description: " & _
    Err.Description
    
End Sub
