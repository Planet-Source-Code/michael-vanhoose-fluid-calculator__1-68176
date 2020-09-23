VERSION 5.00
Begin VB.Form FrmCalculator 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conversion Calculator..."
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbCat 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   2880
      Picture         =   "FrmPics.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "5009"
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2880
      Picture         =   "FrmPics.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "5010"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.TextBox TxtTo 
      Height          =   285
      Left            =   5280
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TxtFrom 
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox CmdTo 
      Height          =   2205
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton CmdConvert 
      Caption         =   "Con&vert"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddCat 
      Height          =   330
      Left            =   2880
      Picture         =   "FrmPics.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "5009"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDeleteCat 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3240
      Picture         =   "FrmPics.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "5010"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.Label LblTo 
      Caption         =   "To:"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label LblFrom 
      Caption         =   "From:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
