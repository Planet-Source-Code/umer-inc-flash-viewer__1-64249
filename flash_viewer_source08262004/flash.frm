VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sean's Flash Viewer v. 1.1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.swf"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sh1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _cx             =   7011
      _cy             =   3836
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu Mnu_load 
         Caption         =   "Load Movie"
         Index           =   1
         Begin VB.Menu mnu_file_load 
            Caption         =   "From File"
            Shortcut        =   ^O
         End
         Begin VB.Menu mnu_URL 
            Caption         =   "From URL"
         End
      End
      Begin VB.Menu mnu_down 
         Caption         =   "Download movie"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_movie 
      Caption         =   "Movie"
      Begin VB.Menu mnu_stop 
         Caption         =   "Pause"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_frame 
         Caption         =   "Next Frame"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_last_frame 
         Caption         =   "Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu Mnu_rewind 
         Caption         =   "Rewind"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_current 
         Caption         =   "Current fame"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_goto 
         Caption         =   "Goto Frame"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_var 
      Caption         =   "Variable"
      Begin VB.Menu mnu_addapt 
         Caption         =   "Addapt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_check 
         Caption         =   "check for variable"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_change 
         Caption         =   "Change Variable"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_Search 
         Caption         =   "Search for variables"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu_Update 
         Caption         =   "Update list"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
      Begin VB.Menu mnu_How 
         Caption         =   "How To Use"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'To store the vresion
Const version = 1.1
'Writen by Sean Oczkowski

'this uses an api call to download a file
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
         ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'Downlaods the file
Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
'a variable used in the download
Dim lngRetVal As Long
   
'downloads the file
lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
'checks if it was done correctly
'if lngRetVal = 0 then the download was good
If lngRetVal = 0 Then DownloadFile = True

End Function

Private Sub form_load()
'disables two menu options
Call change(False)
End Sub

Private Sub Form_Resize()
'moves the flash file
sh1.Top = 0
sh1.Left = 0
'resives it with the from
sh1.Width = Me.ScaleWidth
sh1.Height = Me.ScaleHeight
End Sub

Private Sub change(change As Boolean)
'changes the enabled of the variable menu and the movie menu
'based on the change value
mnu_movie.Enabled = change
mnu_var.Enabled = change
End Sub

Private Sub Form_Unload(Cancel As Integer)
'exit the program when the main menu closes so it stops running
End
End Sub

Private Sub mnu_about_Click()
'shows the info
MsgBox "Created by Umer kk", vbInformation, "About"
End Sub

Private Sub mnu_addapt_Click()
'changes is it is checked or not
mnu_addapt.Checked = Not mnu_addapt.Checked
End Sub

Private Sub mnu_change_Click()
'a temperary value
Dim temp As String
'for the variable enterd
Dim ent As String
'so it does't crash
On Error GoTo wrong2
'get the value of the variable you want to change
ent = InputBox("What value would you like to change", "Variable change")
'adds the search to the search list
adapt (ent)
'gets the new variable value
temp = InputBox("the current value of " & ent & " is " & sh1.GetVariable(ent), ent)

'if the value is not to be changed
If temp <> "" Then
    sh1.SetVariable ent, temp
End If
'exits the sub
Exit Sub
wrong2:
'says it dosen't work
MsgBox "Invalid variable name", vbCritical, "Error"
End Sub

Private Sub mnu_check_Click()
'for the check value
Dim enterd As String
'tell it where to go
On Error GoTo wrong
'get the value enterd
enterd = InputBox("Enter the varable you wish to search for", "Variable search")
'add it to the list
adapt (enterd)
'display the value
MsgBox "The value for " & enterd & " is " & sh1.GetVariable(enterd)
'exit the sub(duh)
Exit Sub
wrong:
'tell them it didn't work
MsgBox "There is no variable called " & enterd, vbCritical, "Sorry"
End Sub

Private Sub mnu_current_Click()
'display the current frame
MsgBox "The current frame is " & sh1.CurrentFrame, , "Current Frame"
End Sub

Private Sub mnu_down_Click()
'for the download file path
Dim temp As String
'get the path to download
temp = InputBox("Enter the URL of the file")
'show the open for the Commem Diolog controll
cd1.ShowSave
'downloads the file and checks if its done
If DownloadFile(temp, cd1.FileName) = True Then
    'says is downloaded
    MsgBox "File loaded", , "Download"
Else
    'says its not downloaded(duh)
    MsgBox "Error", vbCritical, "Error"
End If
End Sub

Private Sub mnu_exit_Click()
'guess
End
End Sub

Private Sub mnu_file_load_Click()
'sets the filter
cd1.Filter = "*.swf"
'opens the commen diolog controll
cd1.ShowOpen
'if there is a file enterd
If cd1.FileName <> "" Then
    'plays movie
    sh1.Movie = cd1.FileName
'enables the menus
Call change(True)
End If
End Sub

Private Sub mnu_frame_Click()
'goes forward one frame
sh1.Forward
End Sub

Private Sub mnu_goto_Click()
'for the frame enterd
Dim ent As String
'tell it where to go when their is an error
On Error GoTo the_mall
'get the frame to goto
ent = InputBox("Enter the frame to goto")
'if they enterd a frame
If ent <> "" Then
    'goto the frame
    sh1.GotoFrame (ent)
End If
'exits the sub
Exit Sub
the_mall:
'tells about the error
MsgBox "Error going to frame", vbCritical, "Error"
End Sub

Private Sub mnu_How_Click()
'show the help form
about.Show
End Sub

Private Sub mnu_last_frame_Click()
'goes back a frame
sh1.Back
End Sub

Private Sub Mnu_rewind_Click()
'rewind
sh1.Rewind
End Sub

Private Sub mnu_Search_Click()
'show the variable form
vars.Show
End Sub

Private Sub mnu_stop_Click()
'change the check value
mnu_stop.Checked = Not mnu_stop.Checked
'if its stoped
If mnu_stop.Checked = False Then
'play
sh1.Play
Else
'stop
sh1.Stop
End If
End Sub

Private Sub mnu_Update_Click()
'for the old searches
Dim olds() As String
'for the new searces
Dim news() As String
'for a control variable
Dim x As Integer
'for a counter
Dim count3 As Integer
'for the new ones to add
Dim add() As String
'for an other control variable
Dim y As Integer
'more counters
Dim count As Integer
Dim count2 As Integer
'for temperanry input
Dim temp As String
'download the file
If DownloadFile("Http://kylesrandomstuff.tripod.com/flash_update.txt", App.Path & "/temp.txt") = True Then
    'make sure the search exist
    If Dir(App.Path & "/search.txt") <> "" Then
        'open thatsearch file
        Open App.Path & "/search.txt" For Input As #1
        'open the temp file
        Open App.Path & "/temp.txt" For Input As #2
            'count the enlements in the search file
            Do Until EOF(1)
                Input #1, temp
                count = count + 1
            Loop
            'redim for the old file
            ReDim olds(1 To count - 1)
            'set count to 0
            count = 0
            'count the elements in the new file
            Do Until EOF(2)
                Input #2, temp
                count = count + 1
            Loop
            'redims for the new file and the things to add
            ReDim news(1 To count)
            ReDim add(1 To count)
        Close #2
        Close #1
 
 'open the files
 Open App.Path & "/search.txt" For Input As #1
 Open App.Path & "/temp.txt" For Input As #2
    'get all the old searces in an arry
    Do Until EOF(1)
        x = x + 1
        Input #1, olds(x)
    Loop
    'set count and reset x
    count = x
    x = 0
    
    Input #2, temp
    'set all the values of the new searches in an array
    Do Until EOF(2)
    x = x + 1
    Input #2, news(x)
    Loop
    
    'set count2 value
    count2 = x
    
    'checks the values to see if there new ones
    For x = 1 To count2
        For y = 1 To count
            If LCase(news(x)) = LCase(olds(y)) Then
            GoTo out1
            End If
        Next y
        'increase to count3 to check how many values need to be added
        count3 = count3 + 1
        'put it in the add array
        add(count3) = news(x)
out1:
        
    Next x
    'close the file
Close #2
Close #1
    
    'if there are values to be added
    If count3 <> 0 Then
    'add them
    Open App.Path & "/search" For Append As #1
        For x = 1 To count3
            'add
            Write #1, add(x)
        Next x
    Close #1
    End If
    
    'open the file and get the number of the lates value
    Open App.Path & "/temp.txt" For Input As #1
        Input #1, temp
    Close #1
    
    'check to see if there is a later version
    If version < temp Then
        MsgBox "There is a later version of this program aviable, Goto kylesrandomstuff.tripod.com for the latest version", vbInformation, "Update"
    End If
    
    'Delete the temp file
    Kill App.Path & "/temp"
    
    Else
    'create the search file with the download version
    Open App.Path & "/search.txt" For Append As #1
    Open App.Path & "/temp.txt" For Input As #2
    Input #2, temp
        Do Until EOF(2)
            Input #2, temp
            Write #2, temp
        Loop
    Close #2
    Close #1
    End If
Else
'say the update failed
MsgBox "Update failed", vbCritical, "Error"
End If
End Sub

Private Sub mnu_URL_Click()
'for the url
Dim temp As String
'tell it what to do when broken
On Error GoTo broke
'get the url
temp = InputBox("Enter the URL of the file")
'play the movie
sh1.Movie = temp
'if the movie works
If temp <> "" Then
    'change the menus
    Call change(True)
End If
'exit the sub
Exit Sub

broke:
'if there was a name enterd
If temp <> "" Then
    'say it was not found
    MsgBox ".Swf not found", , "Error"
End If
End Sub


Private Function adapt(ent As String)
'for a temp string
Dim temp As String
'if adapt is not enabled
If mnu_addapt.Checked = False Then
    Exit Function
End If
'open the file
Open App.Path & "/search.txt" For Input As #1
'go through all the entries
Do Until EOF(1)
    'get the temp string
    Input #1, temp
    'checkes if it exists
    If LCase(temp) = LCase(ent) Then
        'if it does exist in the file then it dose not need to
        'be added
        Close #1
        'this might require some brain power to decover what tis dose
        Exit Function
    End If
Loop
'close the file
Close #1
'open the file
Open App.Path & "/search.txt" For Append As #1
'add the new search
Write #1, ent
'close file
Close #1
End Function
