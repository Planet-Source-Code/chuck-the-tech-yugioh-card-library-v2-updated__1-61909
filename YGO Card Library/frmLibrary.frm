VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibrary 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Yu-Gi-Oh! Card Libray"
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12075
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLibrary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lstSearch 
      Height          =   1110
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   10200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdOpSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   1920
      TabIndex        =   40
      ToolTipText     =   "Open Search Window"
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame fraSearch 
      BorderStyle     =   0  'None
      Caption         =   "Search"
      Height          =   2655
      Left            =   3540
      TabIndex        =   14
      Top             =   10500
      Width           =   7575
      Begin VB.Frame fraCSF 
         Caption         =   "Card Search Filters"
         Height          =   2625
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   7575
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   4200
            TabIndex        =   45
            Text            =   "Level"
            Top             =   1800
            Width           =   495
         End
         Begin VB.ComboBox cmbLevel 
            Height          =   330
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5640
            TabIndex        =   42
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtDesc 
            Height          =   285
            Left            =   4920
            TabIndex        =   29
            Top             =   240
            Width           =   2535
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   375
            Left            =   6600
            TabIndex        =   28
            Top             =   2160
            Width           =   855
         End
         Begin VB.CheckBox chkMonster 
            Caption         =   "Monsters"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.ComboBox cmbAtt 
            Height          =   330
            ItemData        =   "frmLibrary.frx":08CA
            Left            =   120
            List            =   "frmLibrary.frx":08CC
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cmbEffect 
            Height          =   330
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbType 
            Height          =   330
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1080
            Width           =   1455
         End
         Begin VB.ComboBox cmbATK 
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtATK 
            Height          =   285
            Left            =   6120
            TabIndex        =   22
            Text            =   "ATK"
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbDEF 
            Height          =   330
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtDEF 
            Height          =   285
            Left            =   6120
            TabIndex        =   20
            Text            =   "DEF"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkSpell 
            Caption         =   "Spells"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkTrap 
            Caption         =   "Traps"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1950
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox cmbSpell 
            Height          =   330
            Left            =   1440
            TabIndex        =   17
            Text            =   "Spell Type"
            Top             =   1650
            Width           =   1335
         End
         Begin VB.ComboBox cmbTrap 
            Height          =   330
            ItemData        =   "frmLibrary.frx":08CE
            Left            =   1440
            List            =   "frmLibrary.frx":08D0
            TabIndex        =   16
            Text            =   "Trap Type"
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level"
            Height          =   255
            Left            =   3360
            TabIndex        =   43
            Top             =   1560
            Width           =   975
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   0
            X2              =   3240
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   3240
            X2              =   3240
            Y1              =   2520
            Y2              =   1560
         End
         Begin VB.Label lblSrchName 
            Caption         =   "Name Contains"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDesc 
            Caption         =   "Description Contains"
            Height          =   255
            Left            =   3360
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   7440
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblAtt 
            Caption         =   "Attribute"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblEffect 
            Caption         =   "Effect"
            Height          =   255
            Left            =   1800
            TabIndex        =   36
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblSrchType 
            Caption         =   "Monster Type"
            Height          =   255
            Left            =   3360
            TabIndex        =   35
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblATK 
            Caption         =   "Attack"
            Height          =   255
            Left            =   5280
            TabIndex        =   34
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblDEF 
            Caption         =   "DEF"
            Height          =   255
            Left            =   5280
            TabIndex        =   33
            Top             =   1560
            Width           =   975
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   120
            X2              =   3240
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblSpell 
            Caption         =   "Spell Type"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1700
            Width           =   855
         End
         Begin VB.Label lblTrap 
            Caption         =   "Trap Type"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2250
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdTrap 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":08D2
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10680
      Picture         =   "frmLibrary.frx":0EE7
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Display All Trap Cards"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdMagic 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":1588
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9480
      Picture         =   "frmLibrary.frx":1B97
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Display All Magic Cards"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdRitual 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":224B
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8280
      Picture         =   "frmLibrary.frx":2855
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Display All Ritual Monsters"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdFusion 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":2F14
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7080
      Picture         =   "frmLibrary.frx":3547
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Display All Fusion Monsters"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdEffect 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":3BFC
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5880
      Picture         =   "frmLibrary.frx":4222
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Display All Effect Monsters"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdNormal 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":48D9
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4680
      Picture         =   "frmLibrary.frx":4F05
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Display All Normal Monsters"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdAllCard 
      Appearance      =   0  'Flat
      DownPicture     =   "frmLibrary.frx":55AB
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3480
      Picture         =   "frmLibrary.frx":5BBF
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Display All Cards"
      Top             =   7800
      Width           =   1020
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame fraLibrary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7320
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   8280
      Begin VB.Timer tmrArrow 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   12840
         Top             =   2880
      End
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading...."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2400
         TabIndex        =   13
         Top             =   3240
         Width           =   4575
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   54
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   53
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   52
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   51
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   50
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   49
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   48
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   47
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   46
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   6120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   45
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   44
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   43
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   42
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   41
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   40
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   39
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   38
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   37
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   4920
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   36
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   35
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   34
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   33
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   32
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   31
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   30
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   29
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   28
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   3720
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   27
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   26
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   25
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   24
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   23
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   22
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   21
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   20
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   19
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   2520
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   18
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   17
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   16
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   15
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   14
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   13
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   12
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   11
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   10
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   1320
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   9
         Left            =   6840
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   8
         Left            =   6000
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   7
         Left            =   5160
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   6
         Left            =   4320
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   5
         Left            =   3480
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   0
         Left            =   7680
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgArrow 
         Height          =   1590
         Index           =   0
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgArrow 
         Height          =   1575
         Index           =   1
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   5640
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   4
         Left            =   2640
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   3
         Left            =   1800
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   2
         Left            =   960
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1080
         Index           =   1
         Left            =   120
         ToolTipText     =   "Click a card to view details."
         Top             =   120
         Width           =   750
      End
      Begin VB.Image imglibrarybg 
         Appearance      =   0  'Flat
         Height          =   7335
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8295
      End
   End
   Begin RichTextLib.RichTextBox rtbDesc 
      Height          =   2025
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Description of the selected card"
      Top             =   5340
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3572
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmLibrary.frx":6246
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsArrow 
      Left            =   120
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   106
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":62BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":6790
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":6C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":712F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   5055
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label lblATKDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgFrame 
      Height          =   4350
      Left            =   360
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3000
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SCREENWIDTH As Long = 800
Const SCREENHEIGHT As Long = 600


Dim iInc As Integer
Dim strCard As String

'Values for determining when to display scroll arrows
Dim iStart As Integer
Dim iListAll As Integer
Dim iListNor As Integer
Dim iListEff As Integer
Dim iListFus As Integer
Dim iListRit As Integer
Dim iListMag As Integer
Dim iListTrp As Integer
Dim MaxCard As Integer

'used for sorting cards
Dim AllCard As Boolean
Dim Normal As Boolean
Dim Effect As Boolean
Dim Fusion As Boolean
Dim Ritual As Boolean
Dim Magic As Boolean
Dim Trap As Boolean
Dim SearchCard As Boolean
Dim AllCardList() As Integer
Dim NorCardList() As String
Dim EffCardList() As Integer
Dim FusCardList() As Integer
Dim RitCardList() As Integer
Dim MagCardList() As Integer
Dim TrpCardList() As Integer
Dim SearchList() As Integer

Private Sub cmdAllCard_Click()
AllCard = True
Normal = False
Effect = False
Fusion = False
Ritual = False
Magic = False
Trap = False
SearchCard = False
iStart = 0
Load_List
End Sub

Private Sub cmdBack_Click()
End
End Sub

Private Sub cmdEffect_Click()
iStart = 0

    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    Y = 0
    z = 0
    For i = 1 To MaxCard
        For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 1 Then
                EffCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i

AllCard = False
Normal = False
Effect = True
Fusion = False
Ritual = False
Magic = False
Trap = False
SearchCard = False
Load_Effect
End Sub

Private Sub cmdFusion_Click()
iStart = 0

    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    Y = 0
    z = 0
    For i = 1 To MaxCard
       For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 2 Then
                FusCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i


AllCard = False
Normal = False
Effect = False
Fusion = True
Ritual = False
Magic = False
Trap = False
SearchCard = False
Load_Fusion
End Sub

Private Sub cmdMagic_Click()
iStart = 0

    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    Y = 0
    z = 0
    For i = 1 To MaxCard
        For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 5 Then
                MagCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i

AllCard = False
Normal = False
Effect = False
Fusion = False
Ritual = False
Magic = True
Trap = False
SearchCard = False
Load_Magic
End Sub

Private Sub cmdNormal_Click()
iStart = 0
    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    
    For i = 1 To MaxCard
        For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 3 Then
                NorCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i

AllCard = False
Normal = True
Effect = False
Fusion = False
Ritual = False
Magic = False
Trap = False
SearchCard = False
Load_Normal
End Sub

Private Sub cmdReload_Click()
Unload Me
Load Me
End Sub

Private Sub cmdOpSearch_Click()
'Load cmbAtt with Values
cmbAtt.Clear
cmbAtt.List(0) = ""
cmbAtt.List(1) = "Dark"
cmbAtt.List(2) = "Earth"
cmbAtt.List(3) = "Fire"
cmbAtt.List(4) = "Light"
cmbAtt.List(5) = "Water"
cmbAtt.List(6) = "Wind"

'Load cmbEffect with values
cmbEffect.Clear
cmbEffect.List(0) = ""
cmbEffect.List(1) = "Normal"
cmbEffect.List(2) = "Effect"
cmbEffect.List(3) = "Fusion"
cmbEffect.List(4) = "Ritual"
'cmbEffect.List(5) = "Spirit"
'cmbEffect.List(6) = "Toon"
'cmbEffect.List(7) = "Union"
'cmbEffect.List(8) = "Archfiend"

'Load cmbType with Values
cmbType.Clear
cmbType.List(0) = ""
cmbType.List(1) = "Aqua"
cmbType.List(2) = "Beast"
cmbType.List(3) = "Beast-Warrior"
cmbType.List(4) = "Dinosaur"
cmbType.List(5) = "Dragon"
cmbType.List(6) = "Fairy"
cmbType.List(7) = "Fiend"
cmbType.List(8) = "Fish"
cmbType.List(9) = "Insect"
cmbType.List(10) = "Machine"
cmbType.List(11) = "Plant"
cmbType.List(12) = "Pyro"
cmbType.List(13) = "Rock"
cmbType.List(14) = "Reptile"
cmbType.List(15) = "Spellcaster"
cmbType.List(16) = "Sea Serpent"
cmbType.List(17) = "Thunder"
cmbType.List(18) = "Warrior"
cmbType.List(19) = "Winged Beast"
cmbType.List(20) = "Zombie"

'Load cmbLevel with Values
cmbLevel.Clear
cmbLevel.List(0) = "<="
cmbLevel.List(1) = ">="
cmbLevel.List(2) = "="

'Load cmbATK with Values
cmbATK.Clear
cmbATK.List(0) = "<="
cmbATK.List(1) = ">="
cmbATK.List(2) = "="

'Load cmbDEF with Values
cmbDEF.Clear
cmbDEF.List(0) = "<="
cmbDEF.List(1) = ">="
cmbDEF.List(2) = "="

'Load cmbSpell with Values
cmbSpell.Clear
cmbSpell.List(0) = ""
cmbSpell.List(1) = "Normal"
cmbSpell.List(2) = "Continuous"
cmbSpell.List(3) = "Equip"
cmbSpell.List(4) = "Field"
cmbSpell.List(5) = "Quick-Play"
cmbSpell.List(6) = "Ritual"

'Load cmbTrap with Values
cmbTrap.Clear
cmbTrap.List(0) = ""
cmbTrap.List(1) = "Normal"
cmbTrap.List(2) = "Continuous"
cmbTrap.List(3) = "Counter"

'Clear Text Boxes
txtATK.Text = ""
txtDEF.Text = ""
txtName.Text = ""
txtDesc.Text = ""
txtLevel.Text = ""

'Refreshes and puts the Search Frame on top
fraSearch.Top = 700
fraSearch.Top = 200

fraSearch.Refresh
End Sub

Private Sub cmdRitual_Click()
iStart = 0

    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    Y = 0
    z = 0
    For i = 1 To MaxCard
        For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 4 Then
                RitCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i

AllCard = False
Normal = False
Effect = False
Fusion = False
Ritual = True
Magic = False
Trap = False
Load_Ritual
End Sub

Private Sub cmdSearch_Click()
iStart = 0
Dim a As Integer
ReDim SearchList(MaxCard + 1)
On Error Resume Next
If txtLevel <> "" Then If txtLevel < 1 Or txtLevel > 12 Then MsgBox ("Level must be integers between 1 and 12"): Exit Sub
If txtATK <> "" Then If txtATK.Text < 0 Then MsgBox ("ATK and DEF values must be positive integers."): Exit Sub
If txtDEF <> "" Then If txtDEF.Text < 0 Then MsgBox ("ATK and DEF values must be positive integers."): Exit Sub
lstSearch.Clear
For a = 1 To MaxCard
Searchtxt (a)
Next a
SearchCard = True
Load_Search
fraSearch.Top = 700
End Sub

Private Sub cmdTrap_Click()
iStart = 0

    Dim i As Integer
    Dim Y As Integer
    Dim z As Integer
    Y = 0
    z = 0
    For i = 1 To MaxCard
        For z = i + Y To MaxCard
            If Get_Card_Small(z).Frame = 6 Then
                TrpCardList(i) = z
                Exit For
            End If
            Y = Y + 1
        Next z
    Next i

AllCard = False
Normal = False
Effect = False
Fusion = False
Ritual = False
Magic = False
Trap = True
Load_Trap
End Sub

Private Sub cmdCancel_Click()
fraSearch.Top = 700
End Sub

Private Sub Form_Load()

iStart = 0

'sets the max number of cards the library can disply
MaxCard = ReadINI("Cards", "MaxCard", App.Path & "\data\cards.dat")

    'sets screen properties
    Me.ScaleMode = 3 'Pixel
    Me.Width = Screen.TwipsPerPixelX * (SCREENWIDTH)
    Me.Height = Screen.TwipsPerPixelY * (SCREENHEIGHT)
    Me.Show

        'set booleans
    AllCard = True
    Normal = False
    Effect = False
    Fusion = False
    Ritual = False
    Magic = False
    Trap = False
    SearchCard = False
    
    'Set background picture
    Me.Picture = LoadPicture(App.Path & "\pics\bg.jpg")
    
    'Set detail section to default (blank)
    lblATKDEF.Caption = ""
    imgFrame.Picture = LoadPicture(App.Path & "\card_pics\" & "back.jpg")
    imglibrarybg.Picture = LoadPicture(App.Path & "\pics\librarybg.jpg")
    
    'set the arrow pictures
    imgArrow(0) = ilsArrow.ListImages(1).Picture
    imgArrow(1) = ilsArrow.ListImages(2).Picture
    
    'get and set the value for each iList number
    Get_iList
        
    'Re-Dimensions Card Arrays
    ReDim AllCardList(MaxCard)
    ReDim NorCardList(MaxCard)
    ReDim EffCardList(MaxCard)
    ReDim FusCardList(MaxCard)
    ReDim RitCardList(MaxCard)
    ReDim MagCardList(MaxCard)
    ReDim TrpCardList(MaxCard)

Dim i As Integer
    'Creates All Card List
    For i = 1 To MaxCard
        AllCardList(i) = i
    Next i
    
    'Load the All Card list onto the screen
    Load_List
    
    'turn off lblLoading
    lblLoading.Visible = False
    

    
End Sub

Private Sub Get_iList()
Dim vardata As Integer

    vardata = ReadINI("Cards", "iListAll", App.Path & "\data\cards.dat")
    iListAll = vardata
    
    vardata = ReadINI("Cards", "iListNor", App.Path & "\data\cards.dat")
    iListNor = vardata
    
    vardata = ReadINI("Cards", "iListEff", App.Path & "\data\cards.dat")
    iListEff = vardata
    
    vardata = ReadINI("Cards", "iListFus", App.Path & "\data\cards.dat")
    iListFus = vardata
    
    vardata = ReadINI("Cards", "iListRit", App.Path & "\data\cards.dat")
    iListRit = vardata
    
    vardata = ReadINI("Cards", "iListMag", App.Path & "\data\cards.dat")
    iListMag = vardata

    vardata = ReadINI("Cards", "iListTrp", App.Path & "\data\cards.dat")
    iListTrp = vardata

        'iList Let's us know when to stop displaying the down arrow pic.
        'iList values must be multiple of 9
End Sub

Private Sub Load_List()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = AllCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Load_Normal()
Dim i As Integer

On Error Resume Next
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = NorCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Exit Sub
 If Index = 0 Then
    imgArrow(Index) = ilsArrow.ListImages(3).Picture
 Else
    imgArrow(Index) = ilsArrow.ListImages(4).Picture
 End If
 'iInc tells us how many cards to scroll by
 iInc = IIf(Index = 0, -9, 9)
 tmrArrow.Enabled = True
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Index = 0 Then
    imgArrow(Index) = ilsArrow.ListImages(1).Picture
 Else
    imgArrow(Index) = ilsArrow.ListImages(2).Picture
 End If
 'tmrArrow.Enabled = False
End Sub


Private Sub imgCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Exit Sub
If Button = 1 And imgCard(Index).Tag = "" Then Exit Sub
If Button = 1 And imgCard(Index).Picture <> LoadPicture() Then imgCard(Index).BorderStyle = 1: Debug.Print imgCard(Index).Tag: Show_Preview imgCard(Index).Tag: Exit Sub
Show_Preview imgCard(Index).Tag
End Sub

Private Sub imgCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then imgCard(Index).BorderStyle = 0
End Sub

Private Sub tmrArrow_Timer()
iStart = iStart + iInc
    If AllCard Then Load_List: imgArrow(1).Visible = IIf(iStart > iListAll, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListAll Or iStart < 1 Then tmrArrow.Enabled = False
    If Normal Then Load_Normal: imgArrow(1).Visible = IIf(iStart > iListNor, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListNor Or iStart < 1 Then tmrArrow.Enabled = False
    If Effect Then Load_Effect: imgArrow(1).Visible = IIf(iStart > iListEff, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListEff Or iStart < 1 Then tmrArrow.Enabled = False
    If Fusion Then Load_Fusion: imgArrow(1).Visible = IIf(iStart > iListFus, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListFus Or iStart < 1 Then tmrArrow.Enabled = False
    If Ritual Then Load_Ritual: imgArrow(1).Visible = IIf(iStart > iListRit, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListRit Or iStart < 1 Then tmrArrow.Enabled = False
    If Magic Then Load_Magic: imgArrow(1).Visible = IIf(iStart > iListMag, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListMag Or iStart < 1 Then tmrArrow.Enabled = False
    If Trap Then Load_Trap: imgArrow(1).Visible = IIf(iStart > iListTrp, False, True): imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListTrp Or iStart < 1 Then tmrArrow.Enabled = False
    If SearchCard Then Load_Search: imgArrow(1).Visible = IIf(iStart > iListAll, False, True): If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False: imgArrow(0).Visible = IIf(iStart < 1, False, True): If iStart > iListAll Or iStart < 1 Then tmrArrow.Enabled = False
    
    tmrArrow.Enabled = False
End Sub

Private Sub Load_Magic()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = MagCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Load_Effect()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = EffCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Load_Fusion()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = FusCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Load_Ritual()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = RitCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Load_Trap()
Dim i As Integer

On Error Resume Next

  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = TrpCardList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub

Private Sub Show_Preview(strCard As String, Optional strCase As String)
Dim crdCard As Card

  crdCard = Get_Card(strCard)
  
  If crdCard.Name = "" Then Exit Sub
  
  imgFrame.Tag = strCard
  lblName.Caption = " " & crdCard.Name
    If crdCard.Type = "Magic" Or crdCard.Type = "Trap" Then
       lblType.Caption = ""
    Else
        lblType.Caption = " [" & crdCard.Type & "]"
    End If
  rtbDesc = crdCard.Description
  If crdCard.Attack = -1 Then lblATKDEF.Visible = False Else lblATKDEF.Visible = True: lblATKDEF.Caption = "ATK/ " & crdCard.Attack & "   DEF/ " & crdCard.Defence
  
On Error GoTo NoPicture

  imgFrame.Picture = LoadPicture(App.Path & "\card_pics\" & crdCard.Name & ".jpg"): Exit Sub
  
NoPicture: imgFrame.Picture = LoadPicture(App.Path & "\card_pics\nopic.jpg")
End Sub


Private Sub Searchtxt(strCard As Integer)
Dim i As Integer
Dim xyz As Integer
Dim crdCard As Card

crdCard = Get_Card_Small(strCard)

If txtName.Text = "" Then GoTo Desc

For i = 1 To Len(crdCard.Name) - (Len(txtName.Text) - 1)
    If txtName.Text = Mid(crdCard.Name, i, Len(txtName.Text)) Or ChangeCase(txtName.Text) = Mid(crdCard.Name, i, Len(txtName.Text)) Then
    GoTo Desc
    
    End If
    Next i
Exit Sub
Desc:

    If txtDesc.Text = "" Then GoTo CheckRest
    For i = 1 To Len(crdCard.Description) - (Len(txtDesc.Text) - 1)
        If txtDesc.Text = Mid(crdCard.Description, i, Len(txtDesc.Text)) Or ChangeCase(txtDesc.Text) = Mid(crdCard.Description, i, Len(txtDesc.Text)) Then
          GoTo CheckRest
    
    End If
    Next i
Exit Sub
CheckRest:
                'If crdCard.Frame = 6 Then GoTo CheckTrap
                'If crdCard.Frame < 5 Then GoTo CheckMonster
                
If chkSpell.Value = 0 And crdCard.Frame = 5 Then Exit Sub
                'Spell Type Filter
                If cmbSpell.Text = "Normal" And crdCard.Icon <> "Normal" Then Exit Sub
                If cmbSpell.Text = "Normal" And crdCard.Frame = 6 Then Exit Sub
                If cmbSpell.Text = "Equip" And crdCard.Icon <> "Equip" Then Exit Sub
                If cmbSpell.Text = "Field" And crdCard.Icon <> "Field" Then Exit Sub
                'If cmbSpell.Text = "Normal" And cmbTrap.Text <> "Normal" And crdCard.Frame = 6 Then GoTo CheckTrap
                If cmbSpell.Text = "Continuous" And crdCard.Icon <> "Continuous" Then Exit Sub
                If cmbSpell.Text = "Continuous" And crdCard.Frame = 6 Then Exit Sub
                If cmbSpell.Text = "Quick-Play" And crdCard.Icon <> "Quick-Play" Then Exit Sub
                If cmbSpell.Text = "Ritual" And crdCard.Icon <> "Ritual" Then Exit Sub
                
CheckTrap:
            If chkTrap.Value = 0 And crdCard.Frame = 6 Then Exit Sub
                'Trap Type Filter
                If cmbTrap.Text = "Normal" And crdCard.Icon <> "Normal" Then Exit Sub
                
                If cmbTrap.Text = "Continuous" And crdCard.Icon <> "Continuous" Then Exit Sub
                If cmbTrap.Text = "Counter" And crdCard.Icon <> "Counter" Then Exit Sub
            
CheckMonster:
            If chkMonster.Value = 0 And crdCard.Frame < 5 Then Exit Sub
                'Attribute filter
                If cmbAtt.Text = "Dark" And crdCard.Attribute <> 2 Then Exit Sub
                If cmbAtt.Text = "Earth" And crdCard.Attribute <> 3 Then Exit Sub
                If cmbAtt.Text = "Fire" And crdCard.Attribute <> 4 Then Exit Sub
                If cmbAtt.Text = "Light" And crdCard.Attribute <> 5 Then Exit Sub
                If cmbAtt.Text = "Water" And crdCard.Attribute <> 6 Then Exit Sub
                If cmbAtt.Text = "Wind" And crdCard.Attribute <> 7 Then Exit Sub
                
                'Effect Filter
                If cmbEffect.Text = "Normal" And crdCard.Frame <> 3 Then Exit Sub
                If cmbEffect.Text = "Effect" And crdCard.Frame <> 1 Then Exit Sub
                If cmbEffect.Text = "Fusion" And crdCard.Frame <> 2 Then Exit Sub
                If cmbEffect.Text = "Ritual" And crdCard.Frame <> 4 Then Exit Sub
                
                '****ToDo: Add other effect types****
                
                
                'Monter Type Filter
                If cmbType.Text = "Aqua" And crdCard.Type <> "Aqua" Then Exit Sub
                If cmbType.Text = "Beast" And crdCard.Type <> "Beast" Then Exit Sub
                If cmbType.Text = "Beast-Warrior" And crdCard.Type <> "Beast-Warrior" Then Exit Sub
                If cmbType.Text = "Dinosaur" And crdCard.Type <> "Dinosaur" Then Exit Sub
                If cmbType.Text = "Dragon" And crdCard.Type <> "Dragon" Then Exit Sub
                If cmbType.Text = "Fairy" And crdCard.Type <> "Fairy" Then Exit Sub
                If cmbType.Text = "Fiend" And crdCard.Type <> "Fiend" Then Exit Sub
                If cmbType.Text = "Fish" And crdCard.Type <> "Fish" Then Exit Sub
                If cmbType.Text = "Insect" And crdCard.Type <> "Insect" Then Exit Sub
                If cmbType.Text = "Machine" And crdCard.Type <> "Machine" Then Exit Sub
                If cmbType.Text = "Plant" And crdCard.Type <> "Plant" Then Exit Sub
                If cmbType.Text = "Pyro" And crdCard.Type <> "Pyro" Then Exit Sub
                If cmbType.Text = "Rock" And crdCard.Type <> "Rock" Then Exit Sub
                If cmbType.Text = "Reptile" And crdCard.Type <> "Reptile" Then Exit Sub
                If cmbType.Text = "Spellcaster" And crdCard.Type <> "Spellcaster" Then Exit Sub
                If cmbType.Text = "Sea Serpent" And crdCard.Type <> "Sea Serpent" Then Exit Sub
                If cmbType.Text = "Thunder" And crdCard.Type <> "Thunder" Then Exit Sub
                If cmbType.Text = "Warrior" And crdCard.Type <> "Warrior" Then Exit Sub
                If cmbType.Text = "Winged Beast" And crdCard.Type <> "Winged Beast" Then Exit Sub
                If cmbType.Text = "Zombie" And crdCard.Type <> "Zombie" Then Exit Sub
                
On Error Resume Next
                'Level Filter
                If txtLevel.Text = "" Then GoTo LevelTrue
                If cmbLevel.Text = "<=" And crdCard.Level <= txtLevel.Text And crdCard.Level > 0 Then GoTo LevelTrue
                If cmbLevel.Text = ">=" And crdCard.Level >= txtLevel.Text Then GoTo LevelTrue
                If cmbLevel.Text = "=" And crdCard.Level = txtLevel.Text Then GoTo LevelTrue
            Exit Sub
                
LevelTrue:
                'Monster ATK Filter
                If txtATK.Text = "" Then GoTo ATKTrue
                If cmbATK.Text = "<=" And crdCard.Attack <= txtATK.Text And crdCard.Attack >= 0 Then GoTo ATKTrue
                If cmbATK.Text = ">=" And crdCard.Attack >= txtATK.Text Then GoTo ATKTrue
                If cmbATK.Text = "=" And crdCard.Attack = txtATK.Text Then GoTo ATKTrue
            Exit Sub
ATKTrue:
                'Monster DEF Filter
                If txtDEF.Text = "" Then GoTo DEFTrue
                If cmbDEF.Text = "<=" And crdCard.Defence <= txtDEF.Text And crdCard.Attack >= 0 Then GoTo DEFTrue
                If cmbDEF.Text = ">=" And crdCard.Defence >= txtDEF.Text Then GoTo DEFTrue
                If cmbDEF.Text = "=" And crdCard.Defence = txtDEF.Text Then GoTo DEFTrue
            Exit Sub
DEFTrue:
            lstSearch.AddItem crdCard.Name
            
            For xyz = 1 To MaxCard
            If SearchList(xyz) = 0 Then SearchList(xyz) = strCard: Exit For
            Next xyz
            
End Sub


Private Function ChangeCase(txtText As String) As String
Dim strCase As String
Dim intCase As Integer

If Asc(Left(txtText, 1)) > 96 Then intCase = Asc(Left(txtText, 1)) - 32
If Asc(Left(txtText, 1)) < 91 Then intCase = Asc(Left(txtText, 1)) + 32

strCase = Chr(intCase)

ChangeCase = strCase & Right(txtText, Len(txtText) - 1)

End Function

Private Sub Load_Search()
Dim i As Integer

On Error Resume Next
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
  If iStart = 0 Then imgArrow(0).Visible = False: imgArrow(1).Visible = True
  For i% = 1 To 54
    imgCard(i%).Tag = SearchList(iStart + i%)
    imgCard(i%).Picture = LoadPicture()
    imgCard(i%).Picture = LoadPicture(App.Path & "\card_pics\small\" & Get_Card_Small(imgCard(i%).Tag).Name & ".jpg")
    If i% = 54 Then Exit For
    imgCard(i% + 1).Picture = LoadPicture()
  Next i%
  If imgCard(54).Picture = LoadPicture() Then imgArrow(1).Visible = False
End Sub
