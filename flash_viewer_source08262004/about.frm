VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   4530
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Find Info Here"
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label Label2 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Sean's Flash Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'for the help contents
Dim help() As String


Private Sub close_Click()
'hide the form
Me.Hide
End Sub

Private Sub form_load()
'for counting
Dim count As Integer
'for input values
Dim temp As String
'a control variable
Dim x As Integer

'open the help file
Open App.Path & "/help.txt" For Input As #1
'count the elements
Do Until EOF(1)
    Input #1, temp
    count = count + 1
Loop
'close the file
Close #1
'redim based on the elements
ReDim help(1 To 2, 0 To (count) / 2)
'reset count
count = 0
'put the contents of the file in an array
Open App.Path & "/help.txt" For Input As #1
Do Until EOF(1)
    Input #1, help(1, count), help(2, count)
    count = count + 1
Loop
Close #1

'add all the items to the combo box
For x = 0 To count - 1
    cmb.AddItem (help(1, x))
Next x
End Sub

Private Sub Cmb_Click()
'show the selected file
If cmb.ListIndex <> -1 Then
    'display help
    Label2.Caption = help(2, cmb.ListIndex)
End If
End Sub

