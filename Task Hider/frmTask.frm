VERSION 5.00
Begin VB.Form frmTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hides From Windows Task Manager"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   240
      Top             =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Programmed By Billy Conner"
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   1980
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'should work in all NT based systems
'Tested in [Windows XP SP2]

'---------------------
'COMPILE AND RUN!
'---------------------

'---------------------
'Most of this application was Coded by me, Billy Conner
'Credits go out to unknown author of [mSharedMemory.bas] as well.
'---------------------

Option Explicit

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)


'---------------------
'IMPORTANT:  i've noticed that if you call 'ModifyExe' too fast repeatedly, taskmanager stops
'          responding to this app's sendmessage. maybe a security thing? or maybe a flaw.

'NOTE: deleting the item isnt good if the taskmanager isnt paused because it finds it again.
'      when renaming it taskmanager doesnt see a problem so it wont refresh the name.
'      the only time it should refresh is when the user changes the sort order or moves a column

'NOTE: another(a better) way to accomplish this is to subclass the syslistview and
'      listen for the LVM_REDRAWITEMS command or something, then start your searching
'---------------------

'this searches the first 2 items in each column for ".exe" and if found,
'  then it searches each item in the column for our match
Private Function ModifyExe(ByVal pDest As String, ByVal pSource As String, Optional ByVal Delete As Boolean) As Boolean

    Dim RetVal As Long
    Dim RetStr As String
    Dim i As Long, ii As Long
    Dim tCount As Long
    Dim ExitCount As Long
    
    RetVal = FindTaskManager
    '---------------------
    'i would like to get a "column count" instead of seeing if the next column exists..
    '  unfortunately i cant find a way to do that yet.
    'i tried to use columnwidth to test if a column existed(>0). it worked but if the user
    '  resizes a column to 0, then my loop would exit. thus stopping my app from working correctly
    '---------------------
    If RetVal Then
        Do While ii < 26
            RetStr = GetItemText(RetVal, i, ii)
            If RetStr = vbNullString Then ' we've come to the end of the columns
                If i = 0 Then 'was the first loop thru
                    i = 1  'could be the correct column, but .exe not found so add +1
                    ii = -1 'start the column count over
                Else
                    Exit Do
                End If
            ElseIf InStr(LCase$(RetStr), ".exe") Then 'we found the Process column
                tCount = GetItemCount(RetVal)
                For i = i To tCount - 1
                    RetStr = GetItemText(RetVal, i, ii)
                    If LCase$(RetStr) = LCase$(pSource) Then
                        If Delete Then
                            Call DeleteItem(RetVal, i) 'doesnt work as well
                        Else
                            Call SetItemText(RetVal, pDest, i, ii)
                        End If
                        ModifyExe = True
                        
                        '[EXIT DO] can be taken out if the app runs multiple instances
                        'the reason why i put it here is because i am trying to limit
                        ' the amount of unneeded sendmessage calls to taskmanager
                        Exit Do 'should only find 1 instance of itself.
                    End If
                Next i
            End If
            ii = ii + 1
        Loop
    End If

End Function

Private Function FindTaskManager() As Long
    Dim RetVal As Long
    'search for taskmanager (atleast the xp version)
    RetVal = FindWindow("#32770", "Windows Task Manager")
    RetVal = FindWindowEx(RetVal, ByVal 0&, "#32770", vbNullString)
    RetVal = FindWindowEx(RetVal, ByVal 0&, "SysListView32", "Processes")
    
    FindTaskManager = RetVal
    
End Function


Private Sub Form_Load()

    App.TaskVisible = False

End Sub


Private Sub Timer1_Timer()
    If FindTaskManager Then
        Label1.Caption = "Windows Task Manager Found"
        '750 milliseconds doesnt seem too slow to alert taskmanager
        If ModifyExe("explorer.exe", App.EXEName & ".exe") Then
            Label1.Caption = "Item found and changed..."
        End If
        'with some code and creativity, you could also find and change the other info to
        'totally match another process such as [User name] to [SYSTEM]
    Else
        Label1.Caption = "Waiting for Windows Task Manager"
    End If
    
    
End Sub
