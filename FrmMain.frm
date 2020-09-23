VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Flow of Fluids"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCv 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   5880
      TabIndex        =   70
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox TxtKFactor 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   5880
      TabIndex        =   69
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other Tools:"
      Height          =   615
      Left            =   4200
      TabIndex        =   66
      Top             =   4800
      Width           =   3615
      Begin VB.ComboBox CmbTools 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.TextBox TxtArea 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1680
      TabIndex        =   57
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox CmbFluidType 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1560
      TabIndex        =   54
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ComboBox Txtrough 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1800
      TabIndex        =   53
      Text            =   "0.00058"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame FrmType 
      Caption         =   "Formula Type:"
      Height          =   615
      Left            =   4200
      TabIndex        =   51
      Top             =   4080
      Width           =   3615
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.ComboBox TxtDia 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1680
      TabIndex        =   50
      Text            =   "4"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame FrmFrictionFittings 
      Caption         =   "Friction Losses in Fittings:"
      Height          =   2655
      Left            =   3960
      TabIndex        =   26
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   3120
         TabIndex        =   49
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   9
         Left            =   3120
         TabIndex        =   47
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   8
         Left            =   3120
         TabIndex        =   45
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   3120
         TabIndex        =   43
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   3120
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   39
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   37
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   35
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   31
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox CmbFittings 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lbl 
         Caption         =   "Tee Branch:"
         Height          =   255
         Index           =   23
         Left            =   2040
         TabIndex        =   48
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "Tee thru flow:"
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   46
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "180 elbow:"
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   44
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "lr 90 elbow:"
         Height          =   255
         Index           =   20
         Left            =   2040
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "45 elbow:"
         Height          =   255
         Index           =   19
         Left            =   2040
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "90 elbow:"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "Plug Valves:"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "Ball Valves:"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "Angle Valves:"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Lbl 
         Caption         =   "Globe Valves:"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Lbl 
         Caption         =   "Gate Valves:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrmLoos 
      Caption         =   "Friction Losses:"
      Height          =   4335
      Left            =   8040
      TabIndex        =   22
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdDelete 
         Height          =   285
         Left            =   3960
         Picture         =   "FrmMain.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Delete From List"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox TxtLosses 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   63
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox TxtftFittings 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   61
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox TxtTPL 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   59
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   3600
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Diameter:"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lenght:"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Feet of Head:"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox TxtPsi 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtFt 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Lbl 
         Caption         =   "Total Losses:"
         Height          =   255
         Index           =   24
         Left            =   600
         TabIndex        =   64
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Lbl 
         Caption         =   "Total Fitting Losses:"
         Height          =   255
         Index           =   27
         Left            =   600
         TabIndex        =   62
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Lbl 
         Caption         =   "Total Pipe Losses:"
         Height          =   255
         Index           =   26
         Left            =   600
         TabIndex        =   60
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Lbl 
         Caption         =   "Psi:"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Lbl 
         Caption         =   "Feet of Head:"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtDynamic 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2400
      TabIndex        =   21
      Text            =   "60.107"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtAbsolute 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox TxtReynolds 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox TxtFeetPipe 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Text            =   "100"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton CmdCalculate 
      Caption         =   "Calculate..."
      Default         =   -1  'True
      Height          =   375
      Left            =   11160
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox TxtSpecificGravity 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "1.032"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox TxtVelocity 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "10.5"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox TxtFlowRate 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "415"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame FrmComments 
      Caption         =   "Comments:"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   12375
      Begin VB.Label LblComments 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.TextBox TxtKinematic 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Text            =   ".30"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Lbl 
      Caption         =   "Cv Value:"
      Height          =   255
      Index           =   29
      Left            =   4200
      TabIndex        =   71
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "K Factor:"
      Height          =   255
      Index           =   28
      Left            =   4200
      TabIndex        =   68
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Area of Pipe (A=FT):"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   58
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Lbl 
      Caption         =   "weight density of fluid (p=pcft):"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Lbl 
      Caption         =   "Absolute roughness:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "effective roughness:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Reynolds number (R):"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Lbl 
      Caption         =   "Feet of Pipe:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Lbl 
      Caption         =   "specific gravity (lb/ft3):"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Lbl 
      Caption         =   "absolute viscosity (u=centipoise):"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Lbl 
      Caption         =   "Fluild Properties:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Lbl 
      Caption         =   "velocity (ft/sec):"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Lbl 
      Caption         =   "Pipe Diameter (D=in):"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Lbl 
      Caption         =   "Flow Rate (Q=GPM):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbFluidType_Click()
Select Case CmbFluidType
Case "Water @ 60F"
    TxtDynamic = "62.17"
   TxtKinematic = "1.13"
   TxtSpecificGravity = "1.00"
Case "30% propylene glycol water"
    TxtDynamic = "0.672"
   TxtKinematic = ".0003768"
   TxtSpecificGravity = "1.032"
Case "Add it your own"
   TxtDynamic.Enabled = True
   TxtKinematic.Enabled = True
   TxtSpecificGravity.Enabled = True
End Select
End Sub
Private Sub CmbTools_Click()
Select Case CmbTools.Text
Case "Conversion Calculator"
    FrmCalculator.Show 1
End Select
End Sub
Private Sub CmbType_Click()
Dim eachControl As Variant
    For Each eachControl In Me.Controls
        Select Case TypeName(eachControl)
            Case "TextBox":
                eachControl.BackColor = -2147483624
                eachControl.Enabled = False
            End Select
    Next
        TxtDia.Enabled = False
        TxtDia.BackColor = -2147483624
        
        Txtrough.Enabled = False
        Txtrough.BackColor = -2147483624
        
Select Case CmbType.Text
Case "Velocity"
    TxtFlowRate.Enabled = True
    TxtFlowRate.BackColor = &H80000005
    TxtDia.Enabled = True
    TxtDia.BackColor = &H80000005
    LblComments = "Condenser pump suction = 3ft/s " & "Condenser pump discharge = 10ft/s" & vbCr & "Circulation water system 7-12 ft/s " & "General service 5-10ft/s"
      
Case "Capacity"
    TxtArea.Enabled = True
    TxtArea.BackColor = &H80000005
    TxtVelocity.Enabled = True
    TxtVelocity.BackColor = &H80000005
    
Case "Friction Losses in Pipe"
    Txtrough.Enabled = True
    Txtrough.BackColor = &H80000005
    TxtFlowRate.Enabled = True
    TxtFlowRate.BackColor = &H80000005
    TxtVelocity.Enabled = True
    TxtVelocity.BackColor = &H80000005
    TxtDia.Enabled = True
    TxtDia.BackColor = &H80000005
    TxtKinematic.Enabled = True
    TxtKinematic.BackColor = &H80000005
    TxtDynamic.Enabled = True
    TxtDynamic.BackColor = &H80000005
    TxtFeetPipe.Enabled = True
    TxtFeetPipe.BackColor = &H80000005
    TxtAbsolute.Enabled = True
    LblComments = "Darcy-Weisback equation: friction losses in a piping system are a complex function of the system geometry, the fluid properties, and the flow rate in the system."

Case "Pipe Diameter Needed"
    TxtVelocity.Enabled = True
    TxtVelocity.BackColor = &H80000005
    TxtFlowRate.Enabled = True
    TxtFlowRate.BackColor = &H80000005

Case "Friction From K Factor"
    TxtVelocity.Enabled = True
    TxtVelocity.BackColor = &H80000005
    TxtKFactor.Enabled = True
    TxtKFactor.BackColor = &H80000005
Case "Cv coefficient from K"
    TxtKFactor.Enabled = True
    TxtKFactor.BackColor = &H80000005
    LblComments = "Cv is defined as the flow of liquid at 60f in gallons per minute at a pressure drop of one pound per square inch acroos the valve."
Case "K from Cv coefficient"
    TxtCv.Enabled = True
    TxtCv.BackColor = &H80000005
Case "Sudden Contraction 0 < 45 K Factor"

Case "Sudden Enlargement 0 < 45 K Factor"

Case "Horsepower to Raise Water"
    TxtSpecificGravity.Enabled = True
    TxtSpecificGravity.BackColor = &H80000005
    TxtFlowRate.Enabled = True
    TxtFlowRate.BackColor = &H80000005
    TxtFt.Enabled = True
    TxtFt.BackColor = &H80000005
Case "Weight of Water in a Pipe"
    TxtFeetPipe.Enabled = True
    TxtFeetPipe.BackColor = &H80000005
    TxtDia.Enabled = True
    TxtDia.BackColor = &H80000005
   
    

End Select
End Sub

Private Sub CmdAdd_Click()
On Error Resume Next
ListView1.ListItems.Add 1, , TxtDia
ListView1.ListItems(1).ListSubItems.Add , , TxtFeetPipe, , "0.225"
ListView1.ListItems(1).ListSubItems.Add , , TxtFt
End Sub

Private Sub CmdCalculate_Click()
On Error GoTo Err:
Select Case CmbType
Case "Velocity"
    TxtVelocity = Velocitys(TxtFlowRate, TxtDia)
Case "Capacity"
    TxtFlowRate = GetCapacity(TxtArea, TxtVelocity)
Case "Friction Losses in Pipe"
    Call DoDarcy
Case "Pipe Diameter Needed"
    TxtArea = Pipesize(TxtFlowRate, TxtVelocity)
    TxtDia = PipeDia(TxtArea * 144)
Case "Friction From K Factor"
    TxtFt = GetFrictionFromKFactor(TxtVelocity, TxtKFactor)
Case "Cv coefficient from K"
    TxtCv = GetCvFromK(TxtDia, TxtKFactor)
Case "K from Cv coefficient"
    TxtKFactor = GetKFromCv(TxtDia, TxtCv)
Case "Sudden Contraction 0 < 45 K Factor"
    TxtKFactor = GetContractionBig()
Case "Sudden Enlargement 0 < 45 K Factor"
    TxtKFactor = GetEnlargementBig()
Case "Horsepower to Raise Water"
    MsgBox GetHorsepowertoRaiseWater(TxtFlowRate, TxtSpecificGravity, TxtFt) & " Horsepower"
Case "Weight of Water in a Pipe"
    MsgBox GetWeightofWater(TxtFeetPipe, TxtDia) & " LBS."
End Select


Dim i As Integer
Dim value As Double
value = 0
i = 0
For i = 1 To ListView1.ListItems.Count
    'MsgBox ListView1.ListItems(i).ListSubItems(2).Text
        value = value + ListView1.ListItems(i).ListSubItems(2).Text
Next i

TxtTPL = value
  TxtLosses = TxtTPL + TxtftFittings
Exit Sub

Err:
LblComments.Caption = "Warning:" & vbCr & "Missing a value"
End Sub
Private Sub DoDarcy()

TxtReynolds = Reynolds(TxtFlowRate, TxtDynamic, TxtDia, TxtKinematic)
If TxtAbsolute = "" Then
    TxtAbsolute = Colebrook(Txtrough, TxtDia)
Else
    f = TxtAbsolute
    TxtAbsolute = ""
End If
TxtFt = Darcy(TxtFeetPipe, TxtDia, TxtVelocity)

Dim i As Integer
Dim ivalue As Integer
i = 0
TxtftFittings = 0
For i = 0 To 10
    If CmbFittings.Item(i).Text <> "" Then
    Select Case i
    Case 0
        ivalue = 8
    Case 1
        ivalue = 340
    Case 2
        ivalue = 90
    Case 3
        ivalue = 3
    Case 4
        ivalue = 40
    Case 5
        ivalue = 30
    Case 6
        ivalue = 16
    Case 7
        ivalue = 16
    Case 8
        ivalue = 50
    Case 9
        ivalue = 20
    Case 10
        ivalue = 60
    End Select

        TxtftFittings = TxtftFittings + CmbFittings.Item(i).Text * ivalue
    End If
    
   
    
Next i
TxtftFittings = Darcy(TxtftFittings, TxtDia, TxtVelocity)


'TxtftFittings = Darcy(CmbFittings.Item(i).Text * ivalue, TxtDia, TxtVelocity)


End Sub
Private Sub cmdDelete_Click()
If ListView1.ListItems.Count > "0" Then
    Dim Batchlist1count As Integer
        Batchlist1count = ListView1.SelectedItem.Index
            
            ListView1.ListItems.Remove Batchlist1count
Else
    MsgBox "Warning!!! No values In List", vbCritical
End If
End Sub

Private Sub Form_Load()
'CmbFittings
Txtrough.AddItem "0.000005"
Txtrough.AddItem "0.00015"
Txtrough.AddItem "0.0004"
Txtrough.AddItem "0.00058"
Txtrough.AddItem "0.00085"

CmbType.AddItem "Velocity"
CmbType.AddItem "Capacity"
CmbType.AddItem "Pipe Diameter Needed"
CmbType.AddItem "Friction Losses in Pipe"
CmbType.AddItem "Friction From K Factor"
CmbType.AddItem "Cv coefficient from K"
CmbType.AddItem "K from Cv coefficient"
CmbType.AddItem "Sudden Contraction 0 < 45 K Factor"
CmbType.AddItem "Sudden Enlargement 0 < 45 K Factor"
CmbType.AddItem "Horsepower to Raise Water"
CmbType.AddItem "Weight of Water in a Pipe"


Dim i As Integer
i = 0
For i = 0 To 10
    CmbFittings.Item(i).AddItem 1
    CmbFittings.Item(i).AddItem 2
    CmbFittings.Item(i).AddItem 3
    CmbFittings.Item(i).AddItem 4
    CmbFittings.Item(i).AddItem 5
    CmbFittings.Item(i).AddItem 6
    CmbFittings.Item(i).AddItem 7
    CmbFittings.Item(i).AddItem 8
    CmbFittings.Item(i).AddItem 9
    CmbFittings.Item(i).AddItem 10
    CmbFittings.Item(i).AddItem 11
    CmbFittings.Item(i).AddItem 12
Next i
    

    TxtDia.AddItem 0.546
    TxtDia.AddItem 0.742
    TxtDia.AddItem 0.957
    TxtDia.AddItem 1.278
    TxtDia.AddItem 1.5
    TxtDia.AddItem 1.939
    TxtDia.AddItem 2.323
    TxtDia.AddItem 2.9
    TxtDia.AddItem 3.364
    TxtDia.AddItem 3.826
    TxtDia.AddItem 4.813
    TxtDia.AddItem 5.761
    TxtDia.AddItem 7.625
    TxtDia.AddItem 9.75
    TxtDia.AddItem 11.75
    TxtDia.AddItem 13
    TxtDia.AddItem 15
    TxtDia.AddItem 17
    TxtDia.AddItem 19
    TxtDia.AddItem 21

CmbFluidType.AddItem "30% propylene glycol water"
CmbFluidType.AddItem "Water @ 60F"
CmbFluidType.AddItem "Add it your own"

CmbTools.AddItem "Conversion Calculator"

Dim eachControl As Variant
    For Each eachControl In Me.Controls
        Select Case TypeName(eachControl)
            Case "TextBox":
              eachControl.Enabled = False
            End Select
    Next
        Txtrough.Enabled = False
        TxtDia.Enabled = False
End Sub
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then

        ListView1.ListItems.Add 1, , InputBox("Description")
       
                ListView1.ListItems(1).ListSubItems.Add , , InputBox("Description")
         
                        ListView1.ListItems(1).ListSubItems.Add , , InputBox("Feet of Head losses")
    
End If
End Sub
Private Sub TxtDynamic_Change()
On Error Resume Next
TxtSpecificGravity = TxtDynamic / 62.4
End Sub
Private Sub TxtFt_Change()
   On Error Resume Next
    TxtPsi = TxtFt * 0.433463372
End Sub

Private Sub TxtKinematic_Change()
    LblComments = "Kinematic Viscosity x spefific Gravity = Absolute Viscosity" & vbCr & "centistokes x S.G. = Cenitipoise"
End Sub

Private Sub Txtrough_Click()
    LblComments.Caption = "effective roughness of commercial steel pipe = .00058 " & "Steel: new = .00015" & vbCr & "Copper = .000005 " & "Galvanized iron = .0005 " & "Plastic = .000005"
End Sub



