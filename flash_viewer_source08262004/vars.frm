VERSION 5.00
Begin VB.Form vars 
   Caption         =   "Vars"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form2"
   ScaleHeight     =   5025
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "vars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'hide the form
Me.Hide
End Sub

Private Sub Form_Activate()
'call the timer, which updates the list
Call Timer1_Timer
End Sub


Private Sub List1_dblClick()
'the variable to change
Dim ent As String
'for a temperary variable
Dim temp As String
'tell it where to go on error
On Error GoTo problem
'get the variable to change
ent = Left(List1.Text, InStr(1, List1.Text, " ") - 1)
'get the value
temp = Form1.sh1.GetVariable(ent)
'if it didn't work exit the sub
If temp = "" Then
    'guess
    Exit Sub
End If
'get the new value
temp = InputBox("Please enter the new value, the current value of " & ent & " is " & temp, "Change value")
'set the values
Form1.sh1.SetVariable ent, temp
problem:
'refresh the list
Call Timer1_Timer
End Sub

Private Sub Timer1_Timer()
'for a temp value
Dim temp As String
'tell it where to go on an error
On Error Resume Next
'if the form is visable
If Me.Visible = True Then

start:
    'check if the file is there
    If Dir(App.Path & "/search.txt") <> "" Then
        'open the file
        Open App.Path & "/search.txt" For Input As #1
        'clear the file
        vars.List1.Clear
            'go through the file
            Do Until EOF(1)
                'get the search
                Input #1, temp
                'display the value
                vars.List1.AddItem (temp & " = " & Form1.sh1.GetVariable(temp))
            Loop
        Close #1
    Else
        'say the search file was not found
        MsgBox "The search file was not found", vbCritical, "Error"
    End If

End If
End Sub
