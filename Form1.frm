VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   2960
      Width           =   4575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1860
      Left            =   0
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the GetSystemDir API
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim SysDir As String

'the appath
Dim ApPath As String
Private Sub Command1_Click()
DoEvents
'--------------------------------
'check if any files are sellected
'--------------------------------
If List1.SelCount = 0 Then
    MsgBox "No files selected.", vbExclamation, App.Title
Else
    Dim Num As Long
    Num = 0
    '--------------------------------------
    'copy and register the files one-by-one
    '--------------------------------------
    Do Until Num > List1.ListCount - 1
        'check if the file is sellected
        If List1.Selected(Num) = True Then
            'check if the file is already present
            If Dir(SysDir & List1.List(Num), vbNormal) = vbNullString Then
                'copy the file
                FileCopy ApPath & "ocx\" & List1.List(Num), SysDir & List1.List(Num)
            End If
            'register the file
            Shell "regsvr32.exe /s " & SysDir & List1.List(Num)
            List1.List(Num) = List1.List(Num) & vbTab & "- Done."
        End If
        Num = Num + 1
    Loop
    MsgBox "All of selected files were processed successfuly.", vbInformation, App.Title
    'disabling the controls
    Command1.Enabled = False
End If

End Sub
Private Sub Form_Load()
DoEvents
'-----------------------------
'checking for another instance
'-----------------------------
If App.PrevInstance = True Then
    End
End If
'------------------
'finding the appath
'------------------
If Right(App.Path, 1) = "\" Then
    ApPath = App.Path
Else
    ApPath = App.Path & "\"
End If
'----------------------------
'finding the system directory
'----------------------------
Dim Slen As Long
SysDir = Space(255)
Slen = GetSystemDirectory(SysDir, 255)
SysDir = Left(SysDir, Slen)
If Right(SysDir, 1) <> "\" Then
    SysDir = SysDir & "\"
End If
Slen = 0
'---------------------------------------------------
'checking for any OCX files available for extraction
'---------------------------------------------------
If Dir(ApPath & "ocx\*.ocx", vbNormal) = vbNullString Then
    MsgBox "Cannot find any files in the ""\ocx"" directory, to be extracted. " & App.Title & " will now terminate.", vbCritical, App.Title
    End
End If
'------------------------
'configuring the controls
'------------------------
Form1.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & "  by Todor Tzvetkov"
Label1.Caption = "This program will copy and register the following OCX files to your default system directory:" & vbNewLine & SysDir & vbNewLine & "If the files are present, they will not be rewriten, but only registered (via REGSVR32.EXE)."
Command1.Caption = "Extract and Register"
'------------------------------------------
'populate the list1 with the OCX file names
'------------------------------------------
Dim Fname As String
Fname = Dir(ApPath & "ocx\*.ocx", vbNormal)
Do Until Fname = vbNullString
    List1.AddItem Fname
    Fname = Dir$
Loop
'select all the files
Dim Lnum As Integer
For Lnum = 0 To List1.ListCount - 1 Step 1
    List1.Selected(Lnum) = True
Next

End Sub
