VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewAppt 
      Caption         =   "View All Appointments"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12515
      _Version        =   393216
      BackColor       =   15454868
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16776960
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   10
      _Band(0)._MapCol(0)._Name=   "LName"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "FName"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "HospNo"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Street"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "County"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "City"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "Zip"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "Tel"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "DOB"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "Gender"
      _Band(0)._MapCol(9)._RSIndex=   9
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   9000
      TabIndex        =   2
      Top             =   1440
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14677179
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      MonthBackColor  =   -2147483648
      StartOfWeek     =   24510465
      CurrentDate     =   36679
   End
   Begin VB.Label Label15 
      Caption         =   "Click on a date to view appointments for that day"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use the schroll bars to see all data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdViewAppt_Click()
Dim SQL As String
Dim sConnect As String
Dim icol As Integer
Dim cnPatients As ADODB.Connection  'cnPatient is database name
Dim rsAppointment As ADODB.Recordset 'Appointment table
Set cnPatients = New ADODB.Connection
Set rsAppointment = New ADODB.Recordset

sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Data Source=C:\My Documents\Patients97.mdb"

cnPatients.Open (sConnect)
'create sql statement
SQL = "SELECT * FROM Appointment ORDER BY Apptdate, ApptTime "

Set rsAppointment.ActiveConnection = cnPatients
rsAppointment.Open SQL

With MSHFlexGrid1
    .Rows = 1
    .Cols = rsAppointment.Fields.Count
  
    For icol = 0 To rsAppointment.Fields.Count - 1
        .Col = icol
        .Text = rsAppointment.Fields(icol).Name
    Next


    While Not rsAppointment.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment(icol) & ""
        Next
        rsAppointment.MoveNext
    Wend
    .TextMatrix(0, 0) = "Hospital Number"
    .TextMatrix(0, 1) = "Appointment Number"
    .TextMatrix(0, 2) = "First Name"
    .TextMatrix(0, 3) = "Last Name"
    .TextMatrix(0, 4) = "Appointment With"
    .TextMatrix(0, 5) = "Appointment Date"
    .TextMatrix(0, 6) = "Appointment Time"
    .TextMatrix(0, 7) = "Status"
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    '.BackColorFixed = vbWhite
    '.FontWidthFixed = 6
 
    
End With

SizeColumns MSHFlexGrid1

Label16.Caption = "Appointments(All)"


End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

Dim SQL As String
Dim AppDate
Dim sConnect As String
Dim icol As Integer
Dim cnPatients As ADODB.Connection  'cnPatient is database name
Dim rsAppointment As ADODB.Recordset 'Appointment table
Set cnPatients = New ADODB.Connection
Set rsAppointment = New ADODB.Recordset

sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Data Source=C:\My Documents\Patients97.mdb"

cnPatients.Open (sConnect)
'create your sql statement

SQL = "SELECT * From Appointment WHERE ApptDate =#" & SQLDate(DateClicked) & "#" & " ORDER BY ApptDate, ApptTime" & ";"

Set rsAppointment.ActiveConnection = cnPatients
rsAppointment.Open SQL

With MSHFlexGrid1
    .Rows = 1
    .Cols = rsAppointment.Fields.Count


    For icol = 0 To rsAppointment.Fields.Count - 1
        .Col = icol
        .Text = rsAppointment.Fields(icol).Name
    Next


    While Not rsAppointment.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment(icol) & ""
        Next
        rsAppointment.MoveNext
    Wend
End With
Label16.Caption = "Appointments for " & DateClicked

SizeColumns MSHFlexGrid1
End Sub

' Make the FlexGrid's columns big enough to hold all values.
Private Sub SizeColumns(ByVal flx As MSHFlexGrid)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
            wid = TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + 400
    Next c
End Sub

Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "mm/dd/yyyy")
End Function


