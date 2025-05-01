VERSION 5.00
Begin VB.Form frmMainMenu 
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmeConfirmation 
      Height          =   2535
      Left            =   3360
      TabIndex        =   24
      Top             =   4500
      Visible         =   0   'False
      Width           =   15015
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
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
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
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   7305
      End
      Begin VB.Label lblPleaseScanYourId 
         AutoSize        =   -1  'True
         Caption         =   "SCAN your ID for Verification and Order Confirmation"
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
         TabIndex        =   26
         Top             =   240
         Width           =   8640
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
         TabIndex        =   75
         Top             =   1200
         Width           =   2760
      End
      Begin VB.Timer txtTimer 
         Interval        =   1000
         Left            =   17400
         Top             =   360
      End
      Begin VB.Image imgIcon 
         Height          =   1575
         Left            =   4920
         Picture         =   "frmMainMenu.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1680
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   3
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
      TabIndex        =   14
      Top             =   1800
      Width           =   20450
      Begin VB.CommandButton btnOrderList 
         BackColor       =   &H00C0C0C0&
         Caption         =   "View Your Orders"
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
         TabIndex        =   74
         Top             =   240
         Width           =   3000
      End
      Begin VB.CommandButton btnOrderNow 
         BackColor       =   &H0000FF00&
         Caption         =   "SAVE ORDER"
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
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2500
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
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   2500
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
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   3200
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
      Begin VB.Label lblMinusQty4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   10050
         TabIndex        =   85
         Top             =   7465
         Width           =   1200
      End
      Begin VB.Label lblMinusQty2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   3210
         TabIndex        =   84
         Top             =   7465
         Width           =   1200
      End
      Begin VB.Label lblMinusQty5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   16830
         TabIndex        =   83
         Top             =   3780
         Width           =   1200
      End
      Begin VB.Label lblMinusQty3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   10050
         TabIndex        =   82
         Top             =   3780
         Width           =   1200
      End
      Begin VB.Label lblMinusQty1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   3210
         TabIndex        =   81
         Top             =   3780
         Width           =   1200
      End
      Begin VB.Label lblAddQty4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   10050
         TabIndex        =   80
         Top             =   4475
         Width           =   1200
      End
      Begin VB.Label lblAddQty2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   3210
         TabIndex        =   79
         Top             =   4475
         Width           =   1200
      End
      Begin VB.Label lblAddQty5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   16830
         TabIndex        =   78
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label lblAddQty3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   10050
         TabIndex        =   77
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label lblAddQty1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   3210
         TabIndex        =   76
         Top             =   780
         Width           =   1200
      End
      Begin VB.Image imgProduct5 
         Height          =   2655
         Left            =   13880
         Picture         =   "frmMainMenu.frx":15523
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2850
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
         Left            =   18700
         TabIndex        =   67
         Top             =   2800
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtPrice5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   18450
         TabIndex        =   66
         Top             =   1600
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
         Left            =   17100
         TabIndex        =   65
         Top             =   2800
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtQty5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   17080
         TabIndex        =   64
         Top             =   1600
         Width           =   735
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
         Left            =   18015
         TabIndex        =   63
         Top             =   750
         Width           =   2300
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
         TabIndex        =   62
         Top             =   750
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductPrice5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Php 25.00"
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
         Left            =   14115
         TabIndex        =   61
         Top             =   1185
         Width           =   2475
         WordWrap        =   -1  'True
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
         Left            =   11900
         TabIndex        =   60
         Top             =   6520
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtPrice4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   11700
         TabIndex        =   59
         Top             =   5200
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
         Left            =   11900
         TabIndex        =   58
         Top             =   2800
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtPrice3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   11700
         TabIndex        =   57
         Top             =   1600
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
         Left            =   11235
         TabIndex        =   56
         Top             =   4440
         Width           =   2300
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
         Left            =   11235
         TabIndex        =   55
         Top             =   750
         Width           =   2300
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
         Left            =   10250
         TabIndex        =   54
         Top             =   6520
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtQty4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   10260
         TabIndex        =   53
         Top             =   5200
         Width           =   735
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
         Left            =   10250
         TabIndex        =   52
         Top             =   2800
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtQty3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   10260
         TabIndex        =   51
         Top             =   1600
         Width           =   885
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
         TabIndex        =   50
         Top             =   4440
         Width           =   1200
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
         TabIndex        =   49
         Top             =   750
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct4 
         Height          =   2655
         Left            =   7080
         Picture         =   "frmMainMenu.frx":18DA6
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
         TabIndex        =   48
         Top             =   4905
         Width           =   2565
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct3 
         Height          =   2655
         Left            =   7080
         Picture         =   "frmMainMenu.frx":1D1CE
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
         TabIndex        =   47
         Top             =   1185
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct4Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small Tub (150 g.)"
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
         Top             =   4440
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct5Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Smoothie (10 oz.)"
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
         TabIndex        =   45
         Top             =   750
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct3Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small Bowl (150 g.)"
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
         TabIndex        =   44
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
         Left            =   5000
         TabIndex        =   43
         Top             =   6520
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtQty2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   3420
         TabIndex        =   41
         Top             =   5200
         Width           =   795
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
         Left            =   3450
         TabIndex        =   40
         Top             =   6520
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
         TabIndex        =   39
         Top             =   4440
         Width           =   1200
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
         TabIndex        =   38
         Top             =   4900
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct2 
         Height          =   2655
         Left            =   240
         Picture         =   "frmMainMenu.frx":2A881
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   2850
      End
      Begin VB.Label lblProduct2Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Medium (100 g.)"
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
         TabIndex        =   36
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
         Caption         =   "TOTAL"
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
         Left            =   18015
         TabIndex        =   35
         Top             =   300
         Width           =   2300
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
         TabIndex        =   34
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblProductHeader3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SMOOTHIE"
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
         TabIndex        =   33
         Top             =   300
         Width           =   3100
      End
      Begin VB.Label lblPriceHeader2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
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
         Left            =   11235
         TabIndex        =   32
         Top             =   300
         Width           =   2300
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
         TabIndex        =   31
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblProductHeader2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MIX FRUITS"
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
         TabIndex        =   30
         Top             =   300
         Width           =   3100
      End
      Begin VB.Label lblPhp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Php"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   14040
         TabIndex        =   21
         Top             =   7200
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "99,999.99"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   15840
         TabIndex        =   22
         Top             =   6960
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtQty1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   3420
         TabIndex        =   20
         Top             =   1600
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAND TOTAL:"
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
         Left            =   13980
         TabIndex        =   17
         Top             =   6360
         Width           =   2235
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
         Height          =   1365
         Left            =   13980
         TabIndex        =   13
         Top             =   6840
         Width           =   6015
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
         Left            =   5000
         TabIndex        =   11
         Top             =   2800
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
         Left            =   3450
         TabIndex        =   10
         Top             =   2800
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgProduct1 
         Height          =   2655
         Left            =   240
         Picture         =   "frmMainMenu.frx":36FE3
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
         Caption         =   "TOTAL"
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
         Left            =   4395
         TabIndex        =   9
         Top             =   300
         Width           =   2300
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
         Width           =   1200
      End
      Begin VB.Label txtPrice1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   4665
         TabIndex        =   6
         Top             =   1600
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductHeader1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VEGETABLE SALAD"
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
      Begin VB.Label txtPrice2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   4515
         TabIndex        =   8
         Top             =   5200
         Width           =   1980
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
         TabIndex        =   29
         Top             =   750
         Width           =   1200
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
         TabIndex        =   37
         Top             =   1185
         Width           =   2385
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProduct1Pnl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Small (50 g.)"
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
         Left            =   4395
         TabIndex        =   28
         Top             =   750
         Width           =   2300
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
         Left            =   4395
         TabIndex        =   42
         Top             =   4440
         Width           =   2300
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote2 
         Caption         =   "Each products maximum order quantity is 1."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   13920
         TabIndex        =   73
         Top             =   5520
         Width           =   6300
      End
      Begin VB.Label lblNote 
         Caption         =   "NOTE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   13920
         TabIndex        =   72
         Top             =   5040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMainMenu"
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

Private Sub btnBackToMenu_Click()
    If btnCancelOrder.Visible Then
        btnCancelOrder_Click
    End If
    
    End
End Sub

Private Sub btnCancelOrder_Click()
    On Error Resume Next  ' Prevents crashes in case of error
        
    Dim cn As ADODB.Connection
    Set cn = GetConnection()
    
    Dim sqlDelRefId As String
    
    sqlDelRefId = "DELETE FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] WHERE REFID = '" & Trim(txtReferenceId.Caption) & "'"
    cn.Execute sqlDelRefId
    
    cn.Close
    Set cn = Nothing
    
    Reset (True)
    
    On Error GoTo 0  ' Resets error handling
End Sub

Private Sub btnCancelScanning_Click()
    On Error Resume Next

    frmeConfirmation.Visible = False
    txtRegisteredId.Text = ""
    frmTime.Enabled = True
    frmePanel.Enabled = True
    frmeOrderMenu.Enabled = True
    
    On Error GoTo 0
End Sub

Private Sub btnCreateOrder_Click()
    On Error Resume Next
    
    Reset (False)
    
    Dim cn As ADODB.Connection
    Set cn = GetConnection()
    
    Dim rsRdLastRecord As ADODB.Recordset
    Dim sqlLastRecord As String
    Dim sqlInsRefId As String
    
    Dim lastRecord As Integer
    Dim padding As String

    sqlLastRecord = "SELECT MAX([REFID])AS [REFID] FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS]"
    
    Set rsRdLastRecord = New ADODB.Recordset
    rsRdLastRecord.Open sqlLastRecord, cn, adOpenStatic, adLockReadOnly
    
    If rsRdLastRecord.EOF Or IsNull(rsRdLastRecord!RefID) Then
        txtReferenceId.Caption = "FR-" & yearYY & monthMM & "-" & "00001"
    Else
        If Mid(rsRdLastRecord!RefID, 4, 2) = yearYY And Mid(rsRdLastRecord!RefID, 6, 2) = monthMM Then
            lastRecord = Right(rsRdLastRecord!RefID, 5) + 1
            
            padding = String(5 - Len(CStr(lastRecord)), "0") ' Dynamic zero padding
            txtReferenceId.Caption = "FR-" & yearYY & monthMM & "-" & padding & lastRecord
        Else
            txtReferenceId.Caption = "FR-" & yearYY & monthMM & "-" & "00001"
        End If
    End If
    rsRdLastRecord.Close
    Set rsRdLastRecord = Nothing

    sqlInsRefId = "INSERT INTO [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] (REFID) VALUES ('" & txtReferenceId.Caption & "')"
    cn.Execute sqlInsRefId
    
    cn.Close
    Set cn = Nothing
    
    On Error GoTo 0
End Sub

Private Sub btnOrderList_Click()
    On Error Resume Next
    
    frmTime.Enabled = False
    frmePanel.Enabled = False
    frmeOrderMenu.Enabled = False
    frmeConfirmation.Visible = True
    frmeConfirmation.Enabled = True
    txtRegisteredId.SetFocus
    
    On Error GoTo 0
End Sub

Private Sub btnOrderNow_Click()
    On Error Resume Next
    
    Dim hasOrder As Boolean
    hasOrder = False  ' Flag to check if there is at least one ordered product

    ' Check if any quantity is greater than zero
    If val(txtQty1) > 0 Then hasOrder = True
    If val(txtQty2) > 0 Then hasOrder = True
    If val(txtQty3) > 0 Then hasOrder = True
    If val(txtQty4) > 0 Then hasOrder = True
    If val(txtQty5) > 0 Then hasOrder = True

    ' If no product was selected, show error message
    If Not hasOrder Then
        MsgBox "You have not selected any products." & vbNewLine & "Please add products before proceeding.", vbCritical, "No Order Found"
        Exit Sub
    End If
    
    frmTime.Enabled = False
    frmePanel.Enabled = False
    frmeOrderMenu.Enabled = False
    frmeConfirmation.Visible = True
    frmeConfirmation.Enabled = True
    txtRegisteredId.SetFocus
    
    On Error GoTo 0
End Sub

Private Sub Form_Load()
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
        yearYY = Format(rs!ServerTime, "YY")
        monthMM = Format(rs!ServerTime, "MM")
    End If

    ' Cleanup
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    Reset (True)

    On Error GoTo 0
End Sub

Private Sub imgProduct1_Click()
    UpdateProductTotal 1, txtQty1, txtPrice1
End Sub

Private Sub lblAddQty1_Click()
    UpdateProductTotal 1, txtQty1, txtPrice1
End Sub

Private Sub lblAddQty2_Click()
    UpdateProductTotal 2, txtQty2, txtPrice2
End Sub

Private Sub lblAddQty3_Click()
    UpdateProductTotal 3, txtQty3, txtPrice3
End Sub

Private Sub lblAddQty4_Click()
    UpdateProductTotal 4, txtQty4, txtPrice4
End Sub

Private Sub lblAddQty5_Click()
    UpdateProductTotal 5, txtQty5, txtPrice5
End Sub

Private Sub lblMinusQty1_Click()
    UpdateProductTotalMinus 1, txtQty1, txtPrice1
End Sub

Private Sub lblMinusQty2_Click()
    UpdateProductTotalMinus 2, txtQty2, txtPrice2
End Sub

Private Sub lblMinusQty3_Click()
    UpdateProductTotalMinus 3, txtQty3, txtPrice3
End Sub

Private Sub lblMinusQty4_Click()
    UpdateProductTotalMinus 4, txtQty4, txtPrice4
End Sub

Private Sub lblMinusQty5_Click()
    UpdateProductTotalMinus 5, txtQty5, txtPrice5
End Sub

Private Sub lblProduct1Pnl1_Click()
    UpdateProductTotal 1, txtQty1, txtPrice1
End Sub

Private Sub lblProductPrice1_Click()
    UpdateProductTotal 1, txtQty1, txtPrice1
End Sub

Private Sub imgProduct2_Click()
    UpdateProductTotal 2, txtQty2, txtPrice2
End Sub

Private Sub lblProduct2Pnl1_Click()
    UpdateProductTotal 2, txtQty2, txtPrice2
End Sub

Private Sub lblProductPrice2_Click()
    UpdateProductTotal 2, txtQty2, txtPrice2
End Sub

Private Sub imgProduct3_Click()
    UpdateProductTotal 3, txtQty3, txtPrice3
End Sub

Private Sub lblProduct3Pnl1_Click()
    UpdateProductTotal 3, txtQty3, txtPrice3
End Sub

Private Sub lblProductPrice3_Click()
    UpdateProductTotal 3, txtQty3, txtPrice3
End Sub

Private Sub imgProduct4_Click()
    UpdateProductTotal 4, txtQty4, txtPrice4
End Sub

Private Sub lblProduct4Pnl1_Click()
    UpdateProductTotal 4, txtQty4, txtPrice4
End Sub

Private Sub lblProductPrice4_Click()
    UpdateProductTotal 4, txtQty4, txtPrice4
End Sub

Private Sub imgProduct5_Click()
    UpdateProductTotal 5, txtQty5, txtPrice5
End Sub

Private Sub lblProduct5Pnl1_Click()
    UpdateProductTotal 5, txtQty5, txtPrice5
End Sub

Private Sub lblProductPrice5_Click()
    UpdateProductTotal 5, txtQty5, txtPrice5
End Sub

Private Sub txtRegisteredId_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    
    If Not KeyAscii = 13 Then
        Exit Sub
    End If
    
    If Trim(txtRegisteredId.Text) = "" Then
        Exit Sub
    End If
    
    Dim cn As ADODB.Connection
    Set cn = GetConnectionMySql()
    
    Dim rsRdRfid As ADODB.Recordset
    Dim sqlRdRfid As String
    
    sqlRdRfid = "SELECT member_id, emp_no, fname, lname FROM members WHERE rfid_no = '" & Trim(txtRegisteredId.Text) & "'"

    Set rsRdRfid = New ADODB.Recordset
    rsRdRfid.Open sqlRdRfid, cn, adOpenStatic, adLockReadOnly
    
    If rsRdRfid.EOF Or IsNull(rsRdRfid!member_id) Then
        MsgBox "Employee ID is not registered. Please go to HR for registration.", vbCritical, "Registration Error"
        btnCancelScanning_Click
        
        txtRegisteredId.Text = ""
        
        'Cleanup
        rsRdRfid.Close
        Set rsRdRfid = Nothing
        cn.Close
        Set cn = Nothing
        
        Exit Sub
    End If
            
    'Opened thru view orders button
    If btnOrderList.Visible Then
        Me.Hide
        frmOrderViewer.SetId = Trim(txtRegisteredId.Text)
        frmOrderViewer.Show
 
        txtRegisteredId.Text = ""
                
    'Opened thru order creation button
    Else
        'If ID is registered
        Dim dateCheck As String

        Select Case Weekday(Date, vbMonday)
            Case 1 ' Monday
                dateCheck = Date
            Case 2 ' Tuesday
                dateCheck = DateAdd("d", -1, Date)
            Case 3 ' Wednesday
                dateCheck = DateAdd("d", -2, Date)
            Case 4 ' Thursday
                dateCheck = DateAdd("d", -3, Date)
            Case Else
                dateCheck = DateAdd("d", -4, Date)
        End Select

        Dim cnn As ADODB.Connection
        Set cnn = GetConnection()

        Dim rsCheckOrders As ADODB.Recordset
        Dim sqlRdCheckOrders As String
        Dim sqlUpdOrders As String
        Dim sqlInsDetails As String

        sqlRdCheckOrders = "SELECT [DATE], CAST([TIME] AS VARCHAR(5)) AS OrderTime, [REFID] FROM [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] WHERE " & _
                           "[REGISTERED_ID] = '" & Trim(txtRegisteredId.Text) & "' AND " & _
                           "[DATE] BETWEEN '" & Format(dateCheck, "yyyy-MM-dd") & "' AND '" & Format(lblDate, "yyyy-MM-dd") & "' AND " & _
                           "[REFID] <> '" & Trim(txtReferenceId.Caption) & "'"

        Set rsCheckOrders = New ADODB.Recordset
        rsCheckOrders.Open sqlRdCheckOrders, cnn, adOpenDynamic, adLockReadOnly
        
        If rsCheckOrders.EOF Then
            Dim fullName As String
            fullName = Trim(rsRdRfid!fname) & " " & Trim(rsRdRfid!lname)
            ' No orders found, proceed to update using reference ID
            sqlUpdOrders = "UPDATE [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS] SET " & _
                           "[DATE] = '" & Format(lblDate, "yyyy-MM-dd") & "', " & _
                           "[TIME] = '" & Format(lblTime, "HH:mm") & "', " & _
                           "[REGISTERED_ID] = '" & Trim(txtRegisteredId.Text) & "', " & _
                           "[REGISTERED_NAME] = '" & fullName & "', " & _
                           "[TOTAL_PRICE] = '" & txtTotalAmount.Caption & "', " & _
                           "[STATUS] = 'UNCLAIMED' " & _
                           "WHERE REFID = '" & Trim(txtReferenceId.Caption) & "'"

            cnn.Execute sqlUpdOrders

            Dim i As Integer

            For i = 1 To 5
                Dim qtyS As Integer
                Dim priceS As Double

                qtyS = val(Me.Controls("txtQty" & i).Caption)
                priceS = val(Me.Controls("txtPrice" & i).Caption)
                
                If qtyS > 0 Then
                    sqlInsDetails = "INSERT INTO [FREEMEAL].[dbo].[TBL_FRUITS_RECORDS_DETAIL] ([REFID],[PRODUCT_ID],[QTY],[TOTAL_PRICE]) VALUES " & _
                                    "('" & Trim(txtReferenceId.Caption) & "', '" & i & "', '" & qtyS & "', '" & priceS & "')"
                                
                    cnn.Execute sqlInsDetails
                End If
            Next i
            
            MsgBox "Your order submitted successfully." & vbNewLine & "Order Reference ID: " & Trim(txtReferenceId.Caption), vbInformation, "Order Successful"
            
            frmTime.Enabled = True
            frmePanel.Enabled = True
            frmeOrderMenu.Enabled = True
            frmeConfirmation.Visible = False
            frmeConfirmation.Enabled = False
            Reset (True)

        Else
            Dim message As String
            message = "You already submitted your order last " & Format(rsCheckOrders!Date, "MMMM dd, yyyy") & " at " & _
                      Format(rsCheckOrders!OrderTime, "hh:mm AM/PM") & _
                      " with Order Reference ID: " & rsCheckOrders!RefID & vbNewLine & "This order will be cancelled."
            
            MsgBox message, vbCritical, "Existing Order Found"
            
            frmTime.Enabled = True
            frmePanel.Enabled = True
            frmeOrderMenu.Enabled = True
            frmeConfirmation.Visible = False
            frmeConfirmation.Enabled = False
            
            btnCancelOrder_Click
        End If
        
        'Cleanup
        rsCheckOrders.Close
        Set rsCheckOrders = Nothing
        cnn.Close
        Set cnn = Nothing
        
        txtRegisteredId.Text = ""
        
        'Cleanup
        rsRdRfid.Close
        Set rsRdRfid = Nothing
        cn.Close
        Set cn = Nothing
    End If
    
    Exit Sub
    
errHandler:
    MsgBox "Form_Load error: " & Err.Description, vbCritical, "Form Load Failed"
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

Private Sub CheckLabel(txt As Label, lbl As Label)
    If Trim(txt.Caption) = "" Or val(txt.Caption) = 0 Then
        lbl.Visible = False
        txt.Visible = False
    Else
        lbl.Visible = (val(txt.Caption) <> 0)
        txt.Visible = True
    End If
End Sub

Private Sub txtQty1_Change()
    CheckLabel txtQty1, lblPcs1
End Sub

Private Sub txtQty2_Change()
    CheckLabel txtQty2, lblPcs2
End Sub

Private Sub txtQty3_Change()
    CheckLabel txtQty3, lblPcs3
End Sub

Private Sub txtQty4_Change()
    CheckLabel txtQty4, lblPcs4
End Sub

Private Sub txtQty5_Change()
    CheckLabel txtQty5, lblPcs5
End Sub

Private Sub txtPrice1_Change()
    CheckLabel txtPrice1, lblPhp1
End Sub

Private Sub txtPrice2_Change()
    CheckLabel txtPrice2, lblPhp2
End Sub

Private Sub txtPrice3_Change()
    CheckLabel txtPrice3, lblPhp3
End Sub

Private Sub txtPrice4_Change()
    CheckLabel txtPrice4, lblPhp4
End Sub

Private Sub txtPrice5_Change()
    CheckLabel txtPrice5, lblPhp5
End Sub

Private Sub UpdateProductTotal(productId As Integer, txtQty As Object, txtPrice As Object)
    On Error Resume Next

    Dim cn As ADODB.Connection
    Dim rsPrice As ADODB.Recordset
    Dim sqlRdPrice As String
    Dim price As Double

    ' Ensure the quantity doesn't exceed the limit
    If val(txtQty) >= 1 Then
        MsgBox "You have reached the maximum quantity (1) per product." & vbNewLine & "Please try other products.", vbCritical, "Order Limit Reached"
        Exit Sub
    End If
    
    ' Increase the quantity
    txtQty = val(txtQty) + 1

    ' Initialize database connection
    Set cn = GetConnection()
    sqlRdPrice = "SELECT [Price] FROM [FREEMEAL].[dbo].[TBL_FRUITS_PRICELIST] WHERE [ProductId] = " & productId

    ' Execute query
    Set rsPrice = New ADODB.Recordset
    rsPrice.Open sqlRdPrice, cn, adOpenStatic, adLockReadOnly

    ' Check if product has a price
    If rsPrice.EOF Or IsNull(rsPrice!price) Then
        MsgBox "Product does not have a price." & vbNewLine & "Please try other products.", vbCritical, "No Price"
        txtQty = ""
    Else
        price = rsPrice!price
        txtPrice.Caption = Format(price * val(txtQty), "0.00") ' Format with 2 decimal places
    End If

    ' Cleanup
    rsPrice.Close
    cn.Close
    Set rsPrice = Nothing
    Set cn = Nothing

    ' Compute the grand total
    ComputeGrandTotal

    On Error GoTo 0
End Sub

Private Sub UpdateProductTotalMinus(productId As Integer, txtQty As Object, txtPrice As Object)
    On Error Resume Next

    Dim cn As ADODB.Connection
    Dim rsPrice As ADODB.Recordset
    Dim sqlRdPrice As String
    Dim price As Double

    ' Ensure the quantity doesn't exceed the limit
    If val(txtQty) = 0 Then
        Exit Sub
    End If
    
    ' Increase the quantity
    txtQty = val(txtQty) - 1

    ' Initialize database connection
    Set cn = GetConnection()
    sqlRdPrice = "SELECT [Price] FROM [FREEMEAL].[dbo].[TBL_FRUITS_PRICELIST] WHERE [ProductId] = " & productId

    ' Execute query
    Set rsPrice = New ADODB.Recordset
    rsPrice.Open sqlRdPrice, cn, adOpenStatic, adLockReadOnly

    ' Check if product has a price
    If rsPrice.EOF Or IsNull(rsPrice!price) Then
        MsgBox "Product does not have a price." & vbNewLine & "Please try other products.", vbCritical, "No Price"
        txtQty = ""
    Else
        price = rsPrice!price
        txtPrice.Caption = Format(price * val(txtQty), "0.00") ' Format with 2 decimal places
    End If

    ' Cleanup
    rsPrice.Close
    cn.Close
    Set rsPrice = Nothing
    Set cn = Nothing

    ' Compute the grand total
    ComputeGrandTotal

    On Error GoTo 0
End Sub

Private Sub ComputeGrandTotal()
    Dim grandTotal As Double

    ' Add all product totals
    grandTotal = val(txtPrice1.Caption) + val(txtPrice2.Caption) + val(txtPrice3.Caption) + val(txtPrice4.Caption) + val(txtPrice5.Caption)

    ' Display grand total with 2 decimal places
    txtTotalAmount.Caption = Format(grandTotal, "0.00")
End Sub

Public Sub Reset(ByVal isTrue As Boolean)
    If isTrue Then
        txtReferenceId.Caption = ""
        frmeOrderMenu.Enabled = False
        txtTotalAmount.Caption = "0.00"
        btnOrderNow.Visible = False
        btnCancelOrder.Visible = False
        btnCreateOrder.Visible = True
        frmeConfirmation.Visible = False
        btnOrderList.Visible = True
        
        txtQty1.Caption = ""
        txtQty2.Caption = ""
        txtQty3.Caption = ""
        txtQty4.Caption = ""
        txtQty5.Caption = ""
        txtPrice1.Caption = ""
        txtPrice2.Caption = ""
        txtPrice3.Caption = ""
        txtPrice4.Caption = ""
        txtPrice5.Caption = ""
    Else
        frmeOrderMenu.Enabled = True
        txtTotalAmount.Caption = "0.00"
        btnOrderNow.Visible = True
        btnCancelOrder.Visible = True
        btnCreateOrder.Visible = False
        frmeConfirmation.Visible = False
        btnOrderList.Visible = False
    End If

End Sub

Function GetConnection() As ADODB.Connection
    On Error GoTo ERR_HANDLER
        
    Dim cn As New ADODB.Connection
    cn.Open "PROVIDER = MSDASQL;driver={SQL Server};database=FREEMEAL;server=192.168.20.230;uid=sa;pwd=Nbc12#;"
    'cn.Open "PROVIDER = MSDASQL;driver={SQL Server};database=FREEMEAL ;server=NBCP-LT-144;uid=sa;pwd=Nbc12#;"
    Set GetConnection = cn
    
    Exit Function
    
ERR_HANDLER:
    MsgBox "Cannot connect to MS SQL Server", vbCritical, "Database Connection Error"
End Function

Function GetConnectionMySql() As ADODB.Connection
    On Error GoTo ERR_HANDLER
    
    Dim cn As New ADODB.Connection
    cn.Open "Provider=MSDASQL;Driver={MySQL ODBC 8.0 ANSI Driver};Server=192.168.23.64;Database=nbc;User=root;Password=Nbc12#;Option=3;"
    'cn.Open "Provider=MSDASQL;Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;Database=nbc;User=root;Password=Nbc12#;Option=3;"
    Set GetConnectionMySql = cn
    
    Exit Function
    
ERR_HANDLER:
    MsgBox "Connection Error: " & Err.Description, vbCritical, "Database Connection Error"
End Function
