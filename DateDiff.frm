VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   -840
      TabIndex        =   9
      Top             =   1980
      Width           =   4575
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   960
         TabIndex        =   10
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "DateDiff.frx":0000
      Left            =   180
      List            =   "DateDiff.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   1275
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1740
      MaxLength       =   4
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   900
      Width           =   1155
   End
   Begin VB.TextBox D2 
      Height          =   315
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   3
      Top             =   900
      Width           =   315
   End
   Begin VB.TextBox Y1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1740
      MaxLength       =   4
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   420
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1155
   End
   Begin VB.TextBox D1 
      Height          =   315
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   0
      Top             =   420
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Left            =   1740
      TabIndex        =   13
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      Height          =   195
      Left            =   1380
      TabIndex        =   12
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find difference in:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1320
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Dim Date1 As Date
Dim Date2 As Date



Private Sub Command1_Click()
Dim Inter As String
  On Error GoTo Err:
    Date1 = Combo1.Text & " " & D1 & ", " & Y1
    Date2 = Combo2.Text & " " & D2 & ", " & Y2
    
    If Combo3.ListIndex < 0 Then GoTo NoInt
    
    Inter = LCase(Left(Combo3.List(Combo3.ListIndex), 1))
    Label2 = Abs(DateDiff(Inter, Date1, Date2)) & " " & Combo3.List(Combo3.ListIndex)
    
    Exit Sub
    
Err:
    MsgBox "Date Conversion Error."
    Combo2.SetFocus
    Exit Sub
    
NoInt:
    MsgBox "Select Interval."
    Combo3.SetFocus
    Exit Sub
End Sub

Private Sub Form_Load()
    LoadMonths 'Call Sub LoadMonths
    SetTodaysDate 'Set todays date as the first date by default
End Sub



Sub LoadMonths() 'Loads months to Combo Boxes
    Combo1.Clear
    Combo1.AddItem "January"
    Combo1.AddItem "February"
    Combo1.AddItem "March"
    Combo1.AddItem "April"
    Combo1.AddItem "May"
    Combo1.AddItem "June"
    Combo1.AddItem "July"
    Combo1.AddItem "August"
    Combo1.AddItem "September"
    Combo1.AddItem "October"
    Combo1.AddItem "November"
    Combo1.AddItem "December"
    Combo1.ListIndex = 0 'Set to January
    
    Combo2.Clear
    Combo2.AddItem "January"
    Combo2.AddItem "February"
    Combo2.AddItem "March"
    Combo2.AddItem "April"
    Combo2.AddItem "May"
    Combo2.AddItem "June"
    Combo2.AddItem "July"
    Combo2.AddItem "August"
    Combo2.AddItem "September"
    Combo2.AddItem "October"
    Combo2.AddItem "November"
    Combo2.AddItem "December"
    Combo2.ListIndex = 0 'Set to January
End Sub

Function CheckDays(sMonth As String) As Integer
    Select Case LCase$(M)
        Case "january"
            CheckDays = 31
            Exit Function
        Case "february"
            CheckDays = 28
            Exit Function
        Case "march"
            CheckDays = 31
            Exit Function
        Case "april"
            CheckDays = 30
            Exit Function
        Case "may"
            CheckDays = 31
            Exit Function
        Case "june"
            CheckDays = 30
            Exit Function
        Case "july"
            CheckDays = 31
            Exit Function
        Case "august"
            CheckDays = 31
            Exit Function
        Case "spetember"
            CheckDays = 30
            Exit Function
        Case "october"
            CheckDays = 31
            Exit Function
        Case "november"
            CheckDays = 30
            Exit Function
        Case "december"
            CheckDays = 31
            Exit Function
    End Select
End Function

Sub SetTodaysDate()
    Combo1.ListIndex = Month(Date) - 1
    D1 = Day(Date)
    Y1 = Year(Date)
    
    Combo2.ListIndex = -1
    D2 = ""
    Y2 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
            MsgBox "Vote for me on PSC Please!"
            ShellExecute hwnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=9676", vbNullString, vbNullString, SW_SHOW
End Sub
