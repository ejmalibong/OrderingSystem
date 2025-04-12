VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrderViewer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordering System"
   ClientHeight    =   11160
   ClientLeft      =   4335
   ClientTop       =   3135
   ClientWidth     =   20445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   20445
   Begin VB.Frame frmTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20450
      Begin VB.CommandButton btnBackToMenu 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exit Application"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   17520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   2760
      End
      Begin VB.Timer txtTimer 
         Interval        =   1000
         Left            =   15480
         Top             =   720
      End
      Begin MSAdodcLib.Adodc AdodcDetail 
         Height          =   375
         Left            =   120
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmOrderViewer.frx":0000
         OLEDBString     =   $"frmOrderViewer.frx":00D0
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"frmOrderViewer.frx":01A0
         Caption         =   "AdodcHeader"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc AdodcHeader 
         Height          =   375
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmOrderViewer.frx":026C
         OLEDBString     =   $"frmOrderViewer.frx":033C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"frmOrderViewer.frx":040C
         Caption         =   "AdodcHeader"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image imgIcon 
         Height          =   1575
         Left            =   4920
         Picture         =   "frmOrderViewer.frx":0493
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label lblTitle4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IEWER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   930
         Left            =   11760
         TabIndex        =   6
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label lblTitle3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1740
         Left            =   10920
         TabIndex        =   5
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblTitle2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   930
         Left            =   7920
         TabIndex        =   4
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label lblTitle1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1740
         Left            =   6960
         TabIndex        =   3
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17640
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17640
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame frmePanel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   20450
      Begin VB.CommandButton btnSearch 
         BackColor       =   &H00FFFF80&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton btnBack 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back to Main Menu"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   17280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   615
         Left            =   3120
         TabIndex        =   12
         Top             =   260
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   148045825
         CurrentDate     =   45757
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   615
         Left            =   6840
         TabIndex        =   15
         Top             =   255
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   148045825
         CurrentDate     =   45757
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   14
         Top             =   300
         Width           =   405
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblOrderDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1920
      End
   End
   Begin VB.Frame frmeOrderMenu 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8355
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   20450
      Begin MSDataGridLib.DataGrid dgHeader 
         Bindings        =   "frmOrderViewer.frx":159B6
         Height          =   7900
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   13944
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
         DefColWidth     =   1
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ORDER LIST"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "REFID"
            Caption         =   "REFERENCE ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DATE"
            Caption         =   "ORDER DATE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "OrderTime"
            Caption         =   "TIME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "TOTAL_PRICE"
            Caption         =   "TOTAL AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "STATUS"
            Caption         =   "STATUS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgDetail 
         Bindings        =   "frmOrderViewer.frx":159D0
         Height          =   7905
         Left            =   9480
         TabIndex        =   18
         Top             =   300
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   13944
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
         DefColWidth     =   1
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ORDER DETAILS"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ProductName"
            Caption         =   "PRODUCT NAME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "QTY"
            Caption         =   "QUANTITY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "TOTAL_PRICE"
            Caption         =   "AMOUNT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   5999.812
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   3000.189
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOrderViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim yearYY As String
Dim monthMM As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private rfidNumber As String

Public Property Let SetId(val As String)
    rfidNumber = val
End Property

Public Property Get GetId() As String
    GetId = rfidNumber
End Property

Private Sub btnBack_Click()
    Unload Me
    frmMainMenu.Show
    frmMainMenu.Reset (True)
    
    frmMainMenu.frmeConfirmation.Visible = False
    frmMainMenu.txtRegisteredId.Text = ""
    frmMainMenu.frmTime.Enabled = True
    frmMainMenu.frmePanel.Enabled = True
    frmMainMenu.frmeOrderMenu.Enabled = False
End Sub

Private Sub btnBackToMenu_Click()
    End
End Sub

Function GetConnection() As ADODB.Connection
    On Error GoTo ERR_HANDLER
        
    Dim cn As New ADODB.Connection
    cn.Open "Provider=SQLNCLI11;Server=NBCP-LT-144\SQLEXPRESS;Database=FREEMEAL;Uid=sa;Pwd=Nbc12#;"
    Set GetConnection = cn
    
    Exit Function
    
ERR_HANDLER:
    MsgBox "Cannot connect to MS SQL Server", vbCritical, "Database Connection Error"
End Function

Function GetConnectionMySql() As ADODB.Connection
    On Error GoTo ERR_HANDLER
    
    Dim cn As New ADODB.Connection
    cn.Open "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=nbc;User=root;Password=Nbc12#;Option=3;"
    Set GetConnectionMySql = cn
    
    Exit Function
    
ERR_HANDLER:
    MsgBox "Cannot connect to MySQL Server", vbCritical, "Database Connection Error"
End Function

Private Sub btnSearch_Click()
    On Error Resume Next
    
    If Format(dtpFrom.Value, "yyyy-MM-dd") > Format(dtpTo.Value, "yyyy-MM-dd") Then
        MsgBox "Start date is later than to end date.", vbCritical, "Incorrect Date Range"
        dtpFrom.Value = DateAdd("d", -30, Now)
        dtpTo.Value = Now
        Exit Sub
    End If

    Dim sqlRdHeader As String
    
    Dim cn As ADODB.Connection
    Set cn = GetConnection()
    
    sqlRdHeader = "SELECT [REFID], [DATE], CAST([TIME] AS VARCHAR(5)) AS OrderTime, [TOTAL_PRICE], [STATUS] FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] WHERE " & _
                  " CAST([DATE] AS DATE) BETWEEN '" & Format(dtpFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtpTo.Value, "yyyy-MM-dd") & "' AND " & _
                  " [REGISTERED_ID] = '" & GetId & "'" & _
                  " ORDER BY CAST([DATE] AS DATE) ASC, CAST([TIME] AS VARCHAR(5)) ASC "
                
    AdodcHeader.RecordSource = sqlRdHeader
    AdodcHeader.Refresh
    
    Set dgHeader.DataSource = AdodcHeader
    
    If AdodcHeader.Recordset.EOF Then
        MsgBox "No records found", vbCritical, "No Records"
    Else
        AdodcHeader.Recordset.MoveFirst
        dgHeader.Row = 0
        dgHeader.Col = 0
        dgHeader_Click ' simulate click if needed
        dgHeader.SetFocus
    End If
    
    cn.Close
    Set cn = Nothing
    
    On Error GoTo 0
End Sub

Private Sub dgHeader_Click()
    On Error Resume Next
    
    Dim rfidNumber As String
    rfidNumber = dgHeader.Columns(0) ' Assuming that the REFID is in the first column

    LoadDetailData rfidNumber
    
    On Error GoTo 0
End Sub

Private Sub LoadDetailData(rfidNumber As String)
    On Error Resume Next

    Dim rsDetail As ADODB.Recordset
    Dim sqlRdDetail As String
    
    Dim cn As ADODB.Connection
    Set cn = GetConnection()

    ' Query to fetch product details based on REFID
    sqlRdDetail = "SELECT A.REFID, TRIM(B.ProductName) AS ProductName, A.QTY, A.TOTAL_PRICE " & _
                  "FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS_DETAIL] A INNER JOIN [FREEMEAL].[dbo].[TBL_FRUITS_PRICELIST] B " & _
                  "ON A.PRODUCT_ID = B.ProductId WHERE " & _
                  "A.REFID = '" & rfidNumber & "'"
                  
    AdodcDetail.RecordSource = sqlRdDetail
    AdodcDetail.Refresh
    
    Set dgDetail.DataSource = AdodcDetail
    dgDetail.Refresh
    
    ' Clean up
    rsDetail.Close
    Set rsDetail = Nothing
    cn.Close
    Set cn = Nothing

    On Error GoTo 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    dtpFrom.Value = DateAdd("d", -30, Now)
    dtpTo.Value = Now
    
    Dim sqlRdHeader As String
    
    Dim cn As ADODB.Connection
    Set cn = GetConnection()
    
    sqlRdHeader = "SELECT [REFID], [DATE], CAST([TIME] AS VARCHAR(5)) AS OrderTime, [TOTAL_PRICE], [STATUS] FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] WHERE " & _
                  " CAST([DATE] AS DATE) BETWEEN '" & Format(dtpFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtpTo.Value, "yyyy-MM-dd") & "' AND " & _
                  " [REGISTERED_ID] = '" & GetId & "'" & _
                  " ORDER BY CAST([DATE] AS DATE) ASC, CAST([TIME] AS VARCHAR(5)) ASC "
            
    AdodcHeader.RecordSource = sqlRdHeader
    AdodcHeader.Refresh
    
    Set dgHeader.DataSource = AdodcHeader
    
    If Not AdodcHeader.Recordset.EOF Then
        AdodcHeader.Recordset.MoveFirst
        dgHeader.Row = 0
        dgHeader.Col = 0
        dgHeader_Click ' simulate click if needed
        dgHeader.SetFocus
    End If
    
    cn.Close
    Set cn = Nothing
    
    On Error GoTo 0
End Sub

Private Sub txtTimer_Timer()
    On Error Resume Next
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    ' Initialize connection
    Dim cn As ADODB.Connection
    Set cn = GetConnection()

    ' Query to get server date and time
    sql = "SELECT GETDATE() AS ServerTime"

    ' Open recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockReadOnly

    ' Check if data is returned
    If Not rs.EOF Then
        lblDate.Caption = Format(rs!ServerTime, "MMMM dd, yyyy")  ' Ensure lowercase `dd`
        lblTime.Caption = Format(rs!ServerTime, "hh:nn AM/PM")
    End If

    ' Cleanup
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
    On Error GoTo 0
End Sub
