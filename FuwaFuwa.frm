VERSION 5.00
Begin VB.Form frmOrderingSystem 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordering System"
   ClientHeight    =   11160
   ClientLeft      =   4785
   ClientTop       =   2955
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
   Begin VB.Frame frmeConfirmation 
      Height          =   2535
      Left            =   4080
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox txtRegisteredId 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   14775
      End
      Begin VB.CommandButton btnCancelScanning 
         BackColor       =   &H008080FF&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblPleaseTakeNote 
         AutoSize        =   -1  'True
         Caption         =   "Please take note your Order Reference ID for claiming of orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   7305
      End
      Begin VB.Label lblPleaseScanYourId 
         AutoSize        =   -1  'True
         Caption         =   "Please SCAN your ID for Verification and Order Confirmation..."
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
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   10095
      End
   End
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
      Height          =   1900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20450
      Begin VB.CommandButton btnBackToMenu 
         BackColor       =   &H008080FF&
         Caption         =   "BACK TO MAIN MENU"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   16320
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Timer txtTimer 
         Interval        =   1000
         Left            =   17400
         Top             =   360
      End
      Begin VB.Label lblTitle6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   765
         Left            =   3940
         TabIndex        =   76
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblTitle7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Smoothies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   840
         Left            =   2270
         TabIndex        =   75
         Top             =   1005
         Width           =   3705
      End
      Begin VB.Label lblTitle5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mixed  Fruits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   840
         Left            =   1900
         TabIndex        =   74
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblTitle4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YSTEM"
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
         Left            =   12840
         TabIndex        =   73
         Top             =   600
         Width           =   2745
      End
      Begin VB.Label lblTitle3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
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
         Left            =   12000
         TabIndex        =   72
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblTitle2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RDERING"
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
         Left            =   8040
         TabIndex        =   71
         Top             =   600
         Width           =   3675
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
         TabIndex        =   70
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
         TabIndex        =   4
         Top             =   200
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
         TabIndex        =   3
         Top             =   650
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
      TabIndex        =   14
      Top             =   1850
      Width           =   20450
      Begin VB.CommandButton btnOrderNow 
         BackColor       =   &H0000FF00&
         Caption         =   "ORDER NOW"
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
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton btnCancelOrder 
         BackColor       =   &H008080FF&
         Caption         =   "CANCEL"
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
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton btnCreateOrder 
         BackColor       =   &H00FFFF80&
         Caption         =   "CREATE ORDER"
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
         Left            =   10500
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblOrderReferenceId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Reference ID:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   3285
      End
      Begin VB.Label txtReferenceId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FFBO-1907-0000"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   3480
         TabIndex        =   18
         Top             =   265
         Width           =   5300
         WordWrap        =   -1  'True
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
      TabIndex        =   2
      Top             =   2760
      Width           =   20450
      Begin VB.Label lblNote 
         Caption         =   "NOTE: All products maximum order quantity is 2."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   13680
         TabIndex        =   77
         Top             =   5640
         Width           =   6735
      End
      Begin VB.Label lblPhp5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "pesos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   18840
         TabIndex        =   69
         Top             =   2880
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPrice5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   18660
         TabIndex        =   68
         Top             =   1440
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPcs5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pcs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   17280
         TabIndex        =   67
         Top             =   2880
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQty5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   17250
         TabIndex        =   66
         Top             =   1440
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct5Pnl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   18320
         TabIndex        =   65
         Top             =   750
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct5Pnl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   16830
         TabIndex        =   64
         Top             =   750
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductPrice5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 40.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   14160
         TabIndex        =   63
         Top             =   1180
         Width           =   2385
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct5 
         Height          =   2655
         Left            =   13880
         Picture         =   "FuwaFuwa.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2850
      End
      Begin VB.Label lblPhp4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "pesos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12000
         TabIndex        =   62
         Top             =   6600
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPrice4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   11820
         TabIndex        =   61
         Top             =   5040
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPhp3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "pesos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12000
         TabIndex        =   60
         Top             =   2880
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPrice3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   11820
         TabIndex        =   59
         Top             =   1440
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct4Pnl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   11530
         TabIndex        =   58
         Top             =   4440
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct3Pnl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   11530
         TabIndex        =   57
         Top             =   750
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPcs4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pcs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   10440
         TabIndex        =   56
         Top             =   6600
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQty4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   10410
         TabIndex        =   55
         Top             =   5040
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPcs3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pcs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   10440
         TabIndex        =   54
         Top             =   2880
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQty3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   10395
         TabIndex        =   53
         Top             =   1440
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct4Pnl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   10050
         TabIndex        =   52
         Top             =   4440
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct3Pnl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   10050
         TabIndex        =   51
         Top             =   750
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct4 
         Height          =   2655
         Left            =   7080
         Picture         =   "FuwaFuwa.frx":C762
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   2850
      End
      Begin VB.Label lblProductPrice4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 40.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7245
         TabIndex        =   50
         Top             =   4905
         Width           =   2565
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct3 
         Height          =   2655
         Left            =   7080
         Picture         =   "FuwaFuwa.frx":10B8A
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2850
      End
      Begin VB.Label lblProductPrice3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 30.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7245
         TabIndex        =   49
         Top             =   1185
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct4Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small Tub (150g)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   6960
         TabIndex        =   48
         Top             =   4440
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct5Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Smoothie"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   13740
         TabIndex        =   47
         Top             =   750
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct3Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small Bowl (150g)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   6960
         TabIndex        =   46
         Top             =   750
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPhp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "pesos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   45
         Top             =   6600
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQty2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   3555
         TabIndex        =   43
         Top             =   5040
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPcs2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pcs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3600
         TabIndex        =   42
         Top             =   6600
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct2Pnl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3210
         TabIndex        =   41
         Top             =   4440
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductPrice2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 70.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   450
         TabIndex        =   40
         Top             =   4900
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct2 
         Height          =   2655
         Left            =   240
         Picture         =   "FuwaFuwa.frx":1E23D
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   2850
      End
      Begin VB.Label lblProduct2Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Medium Tub (100g)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   120
         TabIndex        =   38
         Top             =   4440
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPriceHeader3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   18320
         TabIndex        =   37
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblQtyHeader3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   16830
         TabIndex        =   36
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblProductHeader3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   13740
         TabIndex        =   35
         Top             =   300
         Width           =   3100
      End
      Begin VB.Label lblPriceHeader2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   11530
         TabIndex        =   34
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblQtyHeader2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   10050
         TabIndex        =   33
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblProductHeader2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6960
         TabIndex        =   32
         Top             =   300
         Width           =   3100
      End
      Begin VB.Label lblPhp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Php"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   16320
         TabIndex        =   21
         Top             =   7320
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "18,888"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   17520
         TabIndex        =   23
         Top             =   7320
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl00 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   19320
         TabIndex        =   22
         Top             =   7320
         Visible         =   0   'False
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQty1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   3570
         TabIndex        =   20
         Top             =   1440
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16200
         TabIndex        =   17
         Top             =   6720
         Width           =   2565
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   16200
         TabIndex        =   13
         Top             =   7200
         Width           =   4095
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPhp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "pesos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   2880
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPcs1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pcs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   2880
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct1 
         Height          =   2655
         Left            =   240
         Picture         =   "FuwaFuwa.frx":2A99F
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2850
      End
      Begin VB.Label lblPriceHeader1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   4690
         TabIndex        =   9
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblQtyHeader1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   3210
         TabIndex        =   7
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblPrice1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   4860
         TabIndex        =   6
         Top             =   1440
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductHeader1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   3100
      End
      Begin VB.Label lblPrice2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   4980
         TabIndex        =   8
         Top             =   5040
         Width           =   1440
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct1Pnl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3210
         TabIndex        =   31
         Top             =   750
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductPrice1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 40.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   465
         TabIndex        =   39
         Top             =   1180
         Width           =   2385
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct1Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small Tub (50g)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   120
         TabIndex        =   12
         Top             =   750
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct1Pnl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   4690
         TabIndex        =   29
         Top             =   750
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct2Pnl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   4690
         TabIndex        =   44
         Top             =   4440
         Width           =   1995
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmOrderingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn              As New ADODB.Connection
Dim rsOperator      As New ADODB.Recordset
Dim rsREFID         As New ADODB.Recordset
Dim rsPrice         As New ADODB.Recordset
Dim rsRegister      As New ADODB.Recordset
Dim rsCounter       As New ADODB.Recordset
Dim rsRegChecker    As New ADODB.Recordset
Dim rsRegExist      As New ADODB.Recordset
Dim sqlQperator     As String
Dim sqlREFID        As String
Dim sqlPrice        As String
Dim sqlCheckPoint   As String
Dim sqlRegister     As String
Dim sqlRegChecker   As String
Dim sqlRegExist     As String
Dim sqlCounterDS    As String
Dim sqlCounterNS    As String
Dim REFID           As String
Dim REFIDlen        As Integer
Dim REFIDnum        As Integer
Dim REFIDzero       As String
Dim REFIDcheckY     As String
Dim REFIDcheckM     As String
Dim YY              As String
Dim MM              As String
Dim ProductID1      As String
Dim ProductID2      As String
Dim ProductID3      As String
Dim ProductID4      As String
Dim Confirm         As String
Dim Today           As String
Dim DateChecker     As String
Dim counter         As Integer
Public Operator     As String
Dim item            As String
Dim xIndex          As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Sub btnBackToMenu_Click()
End
End Sub

Private Sub txtTimer_Timer()
lblDate.Caption = Format(Date, "MMMM DD, YYYY")
lblTime.Caption = Time
End Sub

Private Sub Reset()
frmeOrderMenu.Enabled = False
End Sub
