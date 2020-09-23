VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Load / Save ListBox Demo"
   ClientHeight    =   2670
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5070
   Height          =   3075
   Icon            =   "Form1.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5070
   Top             =   1170
   Width           =   5190
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Selected"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text To Add Here...."
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Mark Gillespie 2001"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
List2.Clear
Dim X!, hf%, tr%, sd%
For loopcount = 0 To List1.ListCount - 1
List1.ListIndex = loopcount
If Not List1.Selected(loopcount) = True Then List2.AddItem List1
Next loopcount

List1.Clear
numnames! = Val(NumOfNames$)
sd% = WritePrivateProfileString("Misc", "NumberOfNames", Trim$(Str$(List2.ListCount)), INIFilename$)
For loopcount = 0 To List2.ListCount - 1
List2.ListIndex = loopcount
tr% = WritePrivateProfileString("Names", Trim$(Str$(loopcount + 1)), Trim$(List2), INIFilename$)

Next loopcount
X = updatelist()

endofroutine:
End Sub


Function updatelist()
On Error Resume Next
List1.Clear
Dim NumOfNames$, df%, numnames!
NumOfNames$ = Space$(128)
df% = GetPrivateProfileString("Misc", "NumberOfNames", "NULL", NumOfNames$, Len(NumOfNames$), INIFilename$)
numnames! = Val(NumOfNames$)
Dim loopcount!, dz%, Name2Add$
For loopcount! = 1 To numnames!
Name2Add$ = Space$(1280)
dz% = GetPrivateProfileString("Names", Str$(Trim(loopcount!)), "NULL", Name2Add$, Len(Name2Add$), INIFilename$)
 
 List1.AddItem Name2Add$
 Next loopcount!
End Function

Private Sub Command3_Click()
X = addnewname(Text1, 0)
End Sub


Function addnewname(addname$, networktype!)
If addname$ = "" Then GoTo abort
If networktype! = 1 Then addname$ = Right$(addname$, Len(addname$) - 2)
NumOfNames$ = Space$(128)
df% = GetPrivateProfileString("Misc", "NumberOfNames", "0", NumOfNames$, Len(NumOfNames$), INIFilename$)
numnames! = Val(NumOfNames$)
sd% = WritePrivateProfileString("Misc", "NumberOfNames", Trim$(Str$(numnames! + 1)), INIFilename$)
sf% = WritePrivateProfileString("Names", Str$(numnames! + 1), Trim$(addname$), INIFilename$)


X = updatelist()
abort:
End Function

Private Sub Form_Activate()
INIFilename$ = ".\" + App.Title + ".ini"
X = updatelist()
End Sub

