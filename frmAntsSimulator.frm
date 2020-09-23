VERSION 5.00
Begin VB.Form frmAntsSimulator 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Ant Farm Simulator"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAntsSimulator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOdor 
      Caption         =   "&Odor Traces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.PictureBox picMapa 
      BackColor       =   &H0091ACB9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Index           =   1
      Left            =   6960
      MousePointer    =   2  'Cross
      ScaleHeight     =   1935
      ScaleWidth      =   1935
      TabIndex        =   7
      Top             =   3480
      Width           =   2000
   End
   Begin VB.PictureBox picTerreno 
      BackColor       =   &H0091ACB9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Index           =   1
      Left            =   350
      ScaleHeight     =   5940
      ScaleWidth      =   5940
      TabIndex        =   6
      Top             =   250
      Width           =   6000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   7680
      Top             =   6360
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "P&ause"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton btnStart 
      BackColor       =   &H00FF8080&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FF8080&
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox picMapa 
      BackColor       =   &H0091ACB9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Index           =   0
      Left            =   6960
      MousePointer    =   2  'Cross
      ScaleHeight     =   1935
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   3480
      Width           =   2000
   End
   Begin VB.PictureBox picTerreno 
      BackColor       =   &H0091ACB9&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Index           =   0
      Left            =   350
      ScaleHeight     =   5940
      ScaleWidth      =   5940
      TabIndex        =   1
      Top             =   250
      Width           =   6000
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   100
      Left            =   6360
      Max             =   129
      SmallChange     =   10
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   350
      Max             =   129
      SmallChange     =   10
      TabIndex        =   5
      Top             =   6250
      Width           =   6000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Population"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
      Begin VB.Label lblPop 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmAntsSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c)Copyright 2002 Roberto Aguirre Maturana
Option Explicit
Const PI = 3.1415926
Const mmTwips = 56.7
Const MaxAnt = 2000
Const MaxTrace = 1000
Const MinTrace = -1000
Dim AntSize As Integer
Dim NestSize As Integer
Public NumAnts As Integer
Dim fx As Double, fy As Double
Dim NestX As Double, NestY As Double
Dim NestFood As Double
Dim GridDimX As Integer, GridDimY As Integer
Dim GridSquare As Double
Dim AntArea As Double
Dim AntZ(MaxAnt) As Double
Dim FoodTrace() As Double
Dim AntsQuad() As Integer
Dim Currenti As Double, Currentj As Double, Futurei As Double, Futurej As Double
Dim Futurei0 As Double, Futurej0 As Double
Dim a As Integer, b As Integer
Dim n As Integer
Dim i As Integer, j As Integer
Dim Odor As Integer
Dim FoodQuad As Integer
Dim FoodRate As Double
Dim MaxBite As Double
Private Type Quad
 Food As Double
 Ants As Integer
 Trace As Double
End Type
Dim Quads() As Quad
Private Type Ant
 X0 As Double
 Y0 As Double
 X1 As Double
 Y1 As Double
 Theta As Double
 Size As Double
 Food As Double
 Life As Double
End Type
Dim Ants(MaxAnt) As Ant
Private Type Food
 i As Integer
 j As Integer
 Seeds As Integer
 Size As Double
End Type
Dim Foods() As Food
Private Function MinTraceQuads(ByVal i As Integer, ByVal j As Integer) As Double
 Dim vi0 As Integer, vj0 As Integer
 Dim vi As Integer, vj As Integer
 Dim MinFood As Integer 'The strongest food trace weaker than current.
 vi0 = 0
 vj0 = 0
 If i > 0 Then
  vi0 = i - 1
 End If
 If j > 0 Then
  vj0 = j - 1
 End If
 vi = vi0
 MinTraceQuads = Quads(i, j).Trace
 MinFood = 0
 While vi <= i + 1 And vi <= GridDimX - 1
  vj = vj0
  While vj <= j + 1 And vj <= GridDimY - 1
   If vi = i Or vj = j Then
    If Quads(vi, vj).Trace = -frmOptions.sliNestTrace.Value Then
     MinTraceQuads = Quads(vi, vj).Trace
     Exit Function
    End If
    If Quads(vi, vj).Trace <= MinTraceQuads And Quads(vi, vj).Trace <> 0 Then
     MinTraceQuads = Quads(vi, vj).Trace
    End If
    If Quads(vi, vj).Trace > MinFood And Quads(vi, vj).Trace < Quads(i, j).Trace Then
      MinFood = Quads(vi, vj).Trace
    End If
   End If
   If vj <= GridDimY - 1 Then
    vj = vj + 1
   End If
  Wend
  If vi <= GridDimX - 1 Then
   vi = vi + 1
  End If
 Wend
 If MinFood > Abs(MinTraceQuads) Then
  MinTraceQuads = MinFood
 End If
 
End Function
Private Function MaxTraceQuads(ByVal i As Integer, ByVal j As Integer) As Double
 Dim vi0 As Integer, vj0 As Integer
 Dim vi As Integer, vj As Integer
 Dim MaxNest As Integer
 vi0 = 0
 vj0 = 0
 If i > 0 Then
  vi0 = i - 1
 End If
 If j > 0 Then
  vj0 = j - 1
 End If
 vi = vi0
 MaxTraceQuads = Quads(i, j).Trace
 MaxNest = 0
 While vi <= i + 1 And vi <= GridDimX - 1
  vj = vj0
  While vj <= j + 1 And vj <= GridDimY - 1
   If vi = i Or vj = j Then
    If Quads(vi, vj).Trace >= MaxTraceQuads Then
     MaxTraceQuads = Quads(vi, vj).Trace
    End If
    If Quads(vi, vj).Trace < MaxNest And Quads(vi, vj).Trace > Quads(i, j).Trace Then
      MaxNest = Quads(vi, vj).Trace
    End If
   End If
   If vj <= GridDimY - 1 Then
    vj = vj + 1
   End If
  Wend
  If vi <= GridDimX - 1 Then
   vi = vi + 1
  End If
 Wend
 If MaxTraceQuads < 0 Then
  MaxTraceQuads = MaxNest
 End If
End Function
Private Sub btnStart_Click()
 Call mnuFileNew_Click
End Sub
Private Sub btnStop_Click()
Timer1.Enabled = Not (Timer1.Enabled)
If (Timer1.Enabled = True) Then
btnStop.Caption = "&Pause"
Else
btnStop.Caption = "&Play"
End If
End Sub

Private Sub btnOdor_Click()
 If Timer1.Enabled Then
  If Odor = 0 Then
   Odor = 1
   btnOdor.Caption = "Hide &Odor Traces"
  Else
   Odor = 0
   btnOdor.Caption = "Show &Odor Traces"
  End If
 End If
End Sub

Private Sub Form_Load()
 'Condiciones iniciales por defecto:
 'Default starting conditions:
 frmOptions.NumAnts.Text = 30
 frmOptions.AntSize.Text = 3
 frmOptions.GridDimX = 30
 frmOptions.GridDimY = frmOptions.GridDimX
 
 frmOptions.sliFoodRate.Max = 100
 frmOptions.sliFoodTrace.Max = MaxTrace
 frmOptions.sliNestTrace.Max = -MinTrace
 
 frmOptions.sliFoodRate.Value = 2
 frmOptions.sliFoodTrace.Value = frmOptions.sliFoodTrace.Max / 2
 frmOptions.sliNestTrace.Value = frmOptions.sliFoodTrace.Value / 2
 
 fx = 10
 fy = 10
 
 HScroll1.Min = fx * picMapa(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
 HScroll1.Max = fx * (picMapa(0).ScaleLeft + picMapa(0).ScaleWidth) - 1 / 2 * picTerreno(0).ScaleWidth
 VScroll1.Min = fy * picMapa(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
 VScroll1.Max = fy * (picMapa(0).ScaleTop + picMapa(0).ScaleHeight) - 1 / 2 * picTerreno(0).ScaleHeight
 HScroll1.Value = 1 / 2 * (HScroll1.Min + HScroll1.Max)
 VScroll1.Value = 1 / 2 * (VScroll1.Min + VScroll1.Max)
End Sub

Private Sub Form_Resize()
Dim twidth As Integer, theight As Integer
If frmAntsSimulator.WindowState <> 1 Then
If frmAntsSimulator.Width < 9465 Then
 frmAntsSimulator.Width = 9465
End If
If frmAntsSimulator.Height < 7545 Then
 frmAntsSimulator.Height = 7545
End If
twidth = picTerreno(0).ScaleWidth
theight = picTerreno(0).ScaleHeight
picTerreno(0).Width = frmAntsSimulator.Width - picTerreno(0).Left - 2800
picTerreno(1).Width = frmAntsSimulator.Width - picTerreno(1).Left - 2800
picTerreno(0).Height = frmAntsSimulator.Height - picTerreno(0).Top - 1100
picTerreno(1).Height = frmAntsSimulator.Height - picTerreno(1).Top - 1100
picMapa(0).Left = frmAntsSimulator.Width - 2400
picMapa(1).Left = frmAntsSimulator.Width - 2400
btnStart.Left = frmAntsSimulator.Width - 2200
btnStop.Left = frmAntsSimulator.Width - 1200
btnOdor.Left = frmAntsSimulator.Width - 2300
'lblPop.Left = frmAntsSimulator.Width - 1700
Frame1.Top = picMapa(0).Top - (3 / 2) * Frame1.Height
Frame1.Left = picMapa(0).Left + (picMapa(0).Width - Frame1.Width) / 2
VScroll1.Height = picTerreno(0).Height
VScroll1.Left = picTerreno(0).Left + picTerreno(0).Width
HScroll1.Width = picTerreno(0).Width
HScroll1.Top = picTerreno(0).Top + picTerreno(0).Height

HScroll1.Min = fx * picMapa(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
HScroll1.Max = fx * (picMapa(0).ScaleLeft + picMapa(0).ScaleWidth) - 1 / 2 * picTerreno(0).ScaleWidth
VScroll1.Min = fy * picMapa(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
VScroll1.Max = fy * (picMapa(0).ScaleTop + picMapa(0).ScaleHeight) - 1 / 2 * picTerreno(0).ScaleHeight


picTerreno(0).ScaleLeft = picTerreno(0).ScaleLeft - 1 / 2 * (picTerreno(0).ScaleWidth - twidth)
picTerreno(0).ScaleTop = picTerreno(0).ScaleTop - 1 / 2 * (picTerreno(0).ScaleHeight - theight)

HScroll1.Value = picTerreno(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
VScroll1.Value = picTerreno(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call mnuFileExit_Click
End Sub

Private Sub HScroll1_Change()
 'picTerreno(0).ScaleLeft = HScroll1.Value - picMapa(0).ScaleWidth * fx / 6
 'picTerreno(0).ScaleLeft = HScroll1.Value - 1 / 4 * (fx * picMapa(0).ScaleWidth - picTerreno(0).ScaleWidth)
 picTerreno(0).ScaleLeft = HScroll1.Value - 1 / 2 * picTerreno(0).ScaleWidth
End Sub

Private Sub mnuFileExit_Click()
 Unload frmOptions
 Unload Me
End Sub
Private Sub mnuFileNew_Click()
 
 Odor = 0
 
 btnStop.Caption = "&Pause"
 btnOdor.Caption = "&Odor Traces"
 HScroll1.Value = (HScroll1.Min + HScroll1.Max) / 2
 VScroll1.Value = (VScroll1.Min + VScroll1.Max) / 2
 If IsNumeric(frmOptions.GridDimX) Then
  GridDimX = CInt(frmOptions.GridDimX)
  GridDimY = CInt(frmOptions.GridDimY)
 End If
 If GridDimX <> Empty Then
  GridSquare = (picMapa(0).ScaleWidth / GridDimX)
 Else
  GridDimX = 20
  GridSquare = (picMapa(0).ScaleWidth / GridDimX)
 End If
 'picMapa(0).Width = picMapa(0).ScaleWidth * 2000 / 130
 'picMapa(0).Height = picMapa(0).ScaleHeight * 2000 / 130
 
 HScroll1.SmallChange = GridSquare * fx
 VScroll1.SmallChange = GridSquare * fy
 HScroll1.LargeChange = picTerreno(0).ScaleWidth / 4
 VScroll1.LargeChange = picTerreno(0).ScaleHeight / 4
 
 'Se crea arreglo para la comida (+ celda auxiliar):
 'A food array is created (+ auxiliary cell):
 ReDim Foods((GridDimX * GridDimY) + 1)
 'ReDim Foodtrace(GridDimX + 1, GridDimY + 1)
 'ReDim AntsQuad(GridDimX + 1, GridDimY + 1)
 ReDim Quads(GridDimX + 1, GridDimY + 1)
 FoodQuad = 0
 If IsNumeric(frmOptions.AntSize.Text) Then
  AntSize = frmOptions.AntSize.Text * mmTwips
  NestSize = GridSquare / 2 * fx
 End If
 If IsNumeric(frmOptions.NumAnts.Text) Then
  NumAnts = frmOptions.NumAnts.Text
 End If
 AntArea = PI * ((AntSize / 6) ^ 2 + (AntSize / 7) ^ 2 + (AntSize / 5) ^ 2)
 
 Randomize Timer
 
 'Se genera el nido:
 'The nest is generated:
 NestX = Int(GridDimX / 2) * GridSquare + GridSquare / 2
 NestY = Int(GridDimY / 2) * GridSquare + GridSquare / 2
 NestFood = 0
 
 'Se genera poblacion original hormigas:
 'Original ants population is generated:
 For a = 0 To NumAnts - 1
  If a = 0 Then
   Ants(a).Size = NestSize
   Ants(a).X0 = NestX
   Ants(a).Y0 = NestY
   Ants(a).X1 = Ants(a).X0 + (-1) ^ (Int(Rnd * 2) + 1) * (Rnd * Ants(a).Size / fx)
   Ants(a).Y1 = Ants(a).Y0 + (-1) ^ (Int(Rnd * 2) + 1) * Sqr((Ants(a).Size / fx) ^ 2 - (Ants(a).X1 - Ants(a).X0) ^ 2)
   Ants(a).Theta = 0
   AntZ(a) = 0
   Ants(a).Food = 0
  Else
   Ants(a).Size = AntSize
   Ants(a).X1 = NestX
   Ants(a).Y1 = NestY
   Ants(a).X0 = Ants(a).X1 + (-1) ^ (Int(Rnd * 2) + 1) * (Rnd * AntSize / fx)
   Ants(a).Y0 = Ants(a).Y1 + (-1) ^ (Int(Rnd * 2) + 1) * Sqr((AntSize / fx) ^ 2 - (Ants(a).X0 - Ants(a).X1) ^ 2)
   AntZ(a) = 0
   Ants(a).Food = 0
  End If
 Next a
 
 'Se generan las fuentes de alimento y su tamaño. Se limpian los rastros de olor:
 'Food sources and its size are generated. Odor traces are resetted:
 FoodRate = 100 - frmOptions.sliFoodRate.Value
 For i = 1 To GridDimX
  For j = 1 To GridDimY
   Quads(i, j).Trace = 0
   Quads(i, j).Ants = 0
   'AntsQuad(i, j) = 0
   If ((i <> (NestX - GridSquare / 2) / GridSquare + 1) Or (j <> (NestY - GridSquare / 2) / GridSquare + 1)) Then
    If Rnd * 1 > FoodRate / 100 Then
     Foods(FoodQuad).i = i
     Foods(FoodQuad).j = j
     Foods(FoodQuad).Seeds = Int(Rnd * 8) + 1
     Foods(FoodQuad).Size = (Rnd * 3 + 1) * 1 / 4 * PI * (GridSquare / 2 * fx) ^ 2
     'Foods(FoodQuad).Size = PI * (GridSquare / 2 * fx) ^ 2
     Quads(Foods(FoodQuad).i, Foods(FoodQuad).j).Food = Foods(FoodQuad).Size
     FoodQuad = FoodQuad + 1
    End If
   End If
  Next j
 Next i
 Quads(Int(GridDimX / 2) + 1, Int(GridDimY / 2) + 1).Ants = NumAnts
 'Nota: FoodQuad en este punto: GridDimX*GridDimY
 'Note: FoodQuad at this point: GridDimX*GridDimY
 'Se inicia la animacion:
 'Animation starts:
 Timer1.Enabled = True
End Sub

Private Sub mnuHelpAbout_Click()
 frmAbout.Show vbModal
End Sub

Private Sub mnuOptions_Click()
 If Not Timer1.Enabled Then
  frmOptions.Show vbModal
 End If
End Sub

Private Sub picMapa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (fx * X > 1 / 2 * picTerreno(0).ScaleWidth) Then
  If (fx * X < fx * picMapa(0).ScaleWidth - 1 / 2 * picTerreno(0).ScaleWidth) Then
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + X)
  Else
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + picMapa(0).ScaleWidth) - 1 / 2 * picTerreno(0).ScaleWidth
  End If
 Else
  HScroll1.Value = fx * picMapa(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
 End If
 If (fy * Y > 1 / 2 * picTerreno(0).ScaleHeight) Then
  If (fy * Y < fy * picMapa(0).ScaleHeight - 1 / 2 * picTerreno(0).ScaleHeight) Then
   VScroll1.Value = fy * (picMapa(0).ScaleTop + Y)
  Else
   VScroll1.Value = fy * (picMapa(0).ScaleTop + picMapa(0).ScaleHeight) - 1 / 2 * picTerreno(0).ScaleHeight
  End If
 Else
  VScroll1.Value = fy * picMapa(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
 End If
End Sub

Private Sub picMapa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
 If (fx * X > 1 / 2 * picTerreno(0).ScaleWidth) Then
  If (fx * X < fx * picMapa(0).ScaleWidth - 1 / 2 * picTerreno(0).ScaleWidth) Then
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + X)
  Else
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + picMapa(0).ScaleWidth) - 1 / 2 * picTerreno(0).ScaleWidth
  End If
 Else
  HScroll1.Value = fx * picMapa(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
 End If
 If (fy * Y > 1 / 2 * picTerreno(0).ScaleHeight) Then
  If (fy * Y < fy * picMapa(0).ScaleHeight - 1 / 2 * picTerreno(0).ScaleHeight) Then
   VScroll1.Value = fy * (picMapa(0).ScaleTop + Y)
  Else
   VScroll1.Value = fy * (picMapa(0).ScaleTop + picMapa(0).ScaleHeight) - 1 / 2 * picTerreno(0).ScaleHeight
  End If
 Else
  VScroll1.Value = fy * picMapa(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
 End If
End If
End Sub
Private Sub picMapa_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (fx * X > 1 / 2 * picTerreno(0).ScaleWidth) Then
  If (fx * X < fx * picMapa(0).ScaleWidth - 1 / 2 * picTerreno(0).ScaleWidth) Then
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + X)
  Else
   HScroll1.Value = fx * (picMapa(0).ScaleLeft + picMapa(0).ScaleWidth) - 1 / 2 * picTerreno(0).ScaleWidth
  End If
 Else
  HScroll1.Value = fx * picMapa(0).ScaleLeft + 1 / 2 * picTerreno(0).ScaleWidth
 End If
 If (fy * Y > 1 / 2 * picTerreno(0).ScaleHeight) Then
  If (fy * Y < fy * picMapa(0).ScaleHeight - 1 / 2 * picTerreno(0).ScaleHeight) Then
   VScroll1.Value = fy * (picMapa(0).ScaleTop + Y)
  Else
   VScroll1.Value = fy * (picMapa(0).ScaleTop + picMapa(0).ScaleHeight) - 1 / 2 * picTerreno(0).ScaleHeight
  End If
 Else
  VScroll1.Value = fy * picMapa(0).ScaleTop + 1 / 2 * picTerreno(0).ScaleHeight
 End If
End Sub

Private Sub Timer1_Timer()

 Dim dx As Double, dy As Double
 Dim AntX00 As Double, AntY00 As Double
 Dim AntX10 As Double, AntY10 As Double
 Dim AntXm As Double, AntYm As Double
 Dim AntXc As Double, AntYc As Double
 Dim TimeNow As Double
 Dim MasMenos As Double
 Dim ni As Integer, nj As Integer
 Dim k As Integer
 Dim FoodCarried As Double
 Randomize Timer
 picMapa(0).Cls
 picTerreno(0).Cls

 'Se recolecta el alimento:
 'The food is gathered:
  b = 0
  MaxBite = PI * (AntSize / 4) ^ 2
  While (b < FoodQuad)
    For a = 0 To NumAnts - 1
     If (Sqr((((((Foods(b).i - 1) * GridSquare) + GridSquare / 2) * fx) - Ants(a).X1 * fx) ^ 2 + (((((Foods(b).j - 1) * GridSquare) + GridSquare / 2) * fy) - Ants(a).Y1 * fy) ^ 2) <= Sqr(Foods(b).Size / PI)) Then
      If Ants(a).Food < MaxBite Then
       If Foods(b).Size > MaxBite Then
        Ants(a).Food = Ants(a).Food + MaxBite
        Foods(b).Size = Foods(b).Size - MaxBite
       Else
        Ants(a).Food = Ants(a).Food + Foods(b).Size
        Foods(b).Size = 0
       End If
      End If
'      If Foods(b).Size > 0 And Abs(Quads(Foods(b).i, Foods(b).j).Trace) < Int(frmOptions.sliFoodTrace.Value * Quads(Foods(b).i, Foods(b).j).Food / (PI * (GridSquare / 2 * fx) ^ 2)) Then
'       Quads(Foods(b).i, Foods(b).j).Trace = Int(frmOptions.sliFoodTrace.Value * Quads(Foods(b).i, Foods(b).j).Food / (PI * (GridSquare / 2 * fx) ^ 2))
'      End If
     End If
    Next a
    'Se borran las zonas donde se acabo el alimento:
    'Out of food zones are erased:
    If Foods(b).Size = 0 Then
     Quads(Foods(b).i, Foods(b).j).Food = 0
     Foods(GridDimX * GridDimY) = Foods(b)
     Foods(b) = Foods(FoodQuad - 1)
     Foods(FoodQuad) = Foods(GridDimX * GridDimY)
     FoodQuad = FoodQuad - 1
    Else
     'El alimento crece y se reproduce:
     'Food grows and breeds:
     'Foods(b).Size = Foods(b).Size + 1 / 100 * MaxBite
     Foods(b).Size = Foods(b).Size + 1 / 5000 * PI * (GridSquare / 2 * fx) ^ 2
     
     If Foods(b).Size >= PI * (GridSquare / 2 * fx) ^ 2 Then
      Foods(b).Size = PI * (GridSquare / 2 * fx) ^ 2
      ni = Foods(b).i + (-1) ^ (Int(Rnd * 2) + 1) * Int(Rnd * 2)
      nj = Foods(b).j + (-1) ^ (Int(Rnd * 2) + 1) * Int(Rnd * 2)
      If ni >= 1 And ni <= GridDimX And nj >= 1 And nj <= GridDimY Then
       If Quads(ni, nj).Ants = 0 And Quads(ni, nj).Food = 0 _
        And Foods(FoodQuad).Size = 0 And Foods(b).Seeds > 0 _
        And (ni <> Foods(b).i Or nj <> Foods(b).j) Then
        FoodQuad = FoodQuad + 1
        Foods(FoodQuad - 1).i = ni
        Foods(FoodQuad - 1).j = nj
        Foods(FoodQuad - 1).Seeds = Int(Rnd * 8) + 1
        Foods(FoodQuad - 1).Size = MaxBite
        'Foods(FoodQuad - 1).Size = 1 / 4 * PI * (GridSquare / 2 * fx) ^ 2
        Quads(ni, nj).Food = Foods(FoodQuad - 1).Size
        Foods(b).Seeds = Foods(b).Seeds - 1
       End If
      End If
     End If
     Quads(Foods(b).i, Foods(b).j).Food = Foods(b).Size
    End If
    If Foods(b).Size > 0 And Abs(Quads(Foods(b).i, Foods(b).j).Trace) < Int(frmOptions.sliFoodTrace.Value * Quads(Foods(b).i, Foods(b).j).Food / (PI * (GridSquare / 2 * fx) ^ 2)) Then
     Quads(Foods(b).i, Foods(b).j).Trace = Int(frmOptions.sliFoodTrace.Value * Quads(Foods(b).i, Foods(b).j).Food / (PI * (GridSquare / 2 * fx) ^ 2))
    End If
    b = b + 1
  Wend
 
 'Se actualizan los rastros:
 'Traces are updated:
 For i = 1 To GridDimX
  For j = 1 To GridDimY
   'Se dibujan los rastros:
   'Traces are drawed:
   If (Quads(i, j).Trace <> 0) Then
    If Quads(i, j).Trace > frmOptions.sliFoodTrace.Value Then
     Quads(i, j).Trace = frmOptions.sliFoodTrace.Value
    End If
    If Quads(i, j).Trace < -frmOptions.sliNestTrace.Value Then
     Quads(i, j).Trace = -frmOptions.sliNestTrace.Value
    End If
    If Quads(i, j).Trace > 0 Then
     picMapa(0).FillColor = RGB(185 + Quads(i, j).Trace * (116 - 185) / frmOptions.sliFoodTrace.Value, 200 + Quads(i, j).Trace * (168 - 200) / frmOptions.sliFoodTrace.Value, 145 + Quads(i, j).Trace * (116 - 145) / frmOptions.sliFoodTrace.Value)
     picTerreno(0).FillColor = RGB(185 + Quads(i, j).Trace * (116 - 185) / frmOptions.sliFoodTrace.Value, 200 + Quads(i, j).Trace * (168 - 200) / frmOptions.sliFoodTrace.Value, 145 + Quads(i, j).Trace * (116 - 145) / frmOptions.sliFoodTrace.Value)
    Else
     picMapa(0).FillColor = RGB(200 + Quads(i, j).Trace * (139 - 200) / -frmOptions.sliNestTrace.Value, 172 + Quads(i, j).Trace * (129 - 172) / -frmOptions.sliNestTrace.Value, 145 + Quads(i, j).Trace * (103 - 145) / -frmOptions.sliNestTrace.Value)
     picTerreno(0).FillColor = RGB(200 + Quads(i, j).Trace * (139 - 200) / -frmOptions.sliNestTrace.Value, 172 + Quads(i, j).Trace * (129 - 172) / -frmOptions.sliNestTrace.Value, 145 + Quads(i, j).Trace * (103 - 145) / -frmOptions.sliNestTrace.Value)
    End If
    'Se dibuja la cantidad de hormigas y la intensidad de los rastros:
    'Ants quantity and trace intensity are drawed:
    If Odor = 1 Then
     picMapa(0).FillStyle = 0
     picMapa(0).Line ((i - 1) * GridSquare, (j - 1) * GridSquare)-(i * GridSquare, j * GridSquare), picMapa(0).FillColor, B
     'If ((i * GridSquare * fx > picTerreno(0).ScaleLeft) And (i * GridSquare * fx < (picTerreno(0).ScaleLeft + picTerreno(0).ScaleWidth)) And (j * GridSquare * fy > picTerreno(0).ScaleTop) And (j * GridSquare * fy < (picTerreno(0).ScaleTop + picTerreno(0).ScaleHeight))) Then
      picTerreno(0).Line ((i - 1) * GridSquare * fx, (j - 1) * GridSquare * fy)-(i * GridSquare * fx, j * GridSquare * fy), RGB(0, 0, 0), B
      picTerreno(0).CurrentX = ((i - 1) * GridSquare + GridSquare / 20) * fx
      picTerreno(0).CurrentY = ((j - 1) * GridSquare + GridSquare / 20) * fy
      If Quads(i, j).Food = 0 Then
       If Quads(i, j).Trace > 0 Then
        picTerreno(0).Print Int(Quads(i, j).Trace)
        'picTerreno(0).Print Int(Quads(i, j).FoodTrace)
       Else
        If Quads(i, j).Trace < 0 Then
         picTerreno(0).Print Int(-1 * Quads(i, j).Trace)
         'picTerreno(0).Print Int(-1 * Quads(i, j).NestTrace)
        End If
       End If
      End If
      picTerreno(0).CurrentX = ((i - 1) * GridSquare + GridSquare * 14 / 20) * fx
      picTerreno(0).CurrentY = ((j - 1) * GridSquare + GridSquare * 15 / 20) * fy
      picTerreno(0).Print Quads(i, j).Ants
      'picTerreno(0).Print AntsQuad(i, j)
     'End If
    End If
     If Quads(i, j).Trace > 0 Then
      Quads(i, j).Trace = Quads(i, j).Trace - 1
     Else
      If (Quads(i, j).Trace < 0) _
         And (((i - 1) * GridSquare) + GridSquare / 2 <> NestX) Or (((j - 1) * GridSquare) + GridSquare / 2 <> NestY) Then
       Quads(i, j).Trace = Quads(i, j).Trace + 1
      End If
     End If
   End If
  Next j
 Next i
 
'Se dibujan las fuentes de alimento:
'Food sources are drawed:
 For b = 0 To FoodQuad - 1
  If Foods(b).Size > 0 Then
   picMapa(0).FillColor = RGB(16, 152, 16)
   picMapa(0).FillStyle = 0
   picMapa(0).Circle ((((Foods(b).i - 1) * GridSquare) + GridSquare / 2), (((Foods(b).j - 1) * GridSquare) + GridSquare / 2)), Sqr(Foods(b).Size / PI) / fx, RGB(0, 0, 0)
    'If ((i * GridSquare * fx > picTerreno(0).ScaleLeft) And (i * GridSquare * fx < (picTerreno(0).ScaleLeft + picTerreno(0).ScaleWidth)) And (j * GridSquare * fy > picTerreno(0).ScaleTop) And (j * GridSquare * fy < (picTerreno(0).ScaleTop + picTerreno(0).ScaleHeight))) Then
     picTerreno(0).FillColor = RGB(16, 152, 16)
     'Call Ellipse(picTerreno(0).hdc, FoodPosi * fx - Sqr(FoodCoord(i, j) / PI), FoodPosj * fy - Sqr(FoodCoord(i, j) / PI), FoodPosi * fx + Sqr(FoodCoord(i, j) / PI), FoodPosj * fy + Sqr(FoodCoord(i, j) / PI))
     picTerreno(0).Circle ((((Foods(b).i - 1) * GridSquare) + GridSquare / 2) * fx, (((Foods(b).j - 1) * GridSquare) + GridSquare / 2) * fy), Sqr(Foods(b).Size / PI), RGB(0, 0, 0)
     'picTerreno(0).CurrentX = (((Foods(b).i - 1) * GridSquare) + GridSquare / 2) * fx
     'picTerreno(0).CurrentY = (((Foods(b).j - 1) * GridSquare) + GridSquare / 2) * fy
     'picTerreno(0).ForeColor = RGB(255, 255, 0)
     'picTerreno(0).Print Int(Foods(b).Size)
     'picTerreno(0).ForeColor = RGB(0, 0, 0)
    'End If
    'picMapa(0).FillStyle = 1
    If Odor = 1 Then
      picTerreno(0).ForeColor = RGB(255, 255, 0)
      picTerreno(0).CurrentX = ((Foods(b).i - 1) * GridSquare + GridSquare / 20) * fx
      picTerreno(0).CurrentY = ((Foods(b).j - 1) * GridSquare + GridSquare / 20) * fy
      If Quads(Foods(b).i, Foods(b).j).Trace > 0 Then
       picTerreno(0).Print Int(Quads(Foods(b).i, Foods(b).j).Trace)
       'picTerreno(0).Print Int(Quads(i, j).FoodTrace)
      Else
       If Quads(Foods(b).i, Foods(b).j).Trace < 0 Then
        picTerreno(0).Print -Int(Quads(Foods(b).i, Foods(b).j).Trace)
        'picTerreno(0).Print Int(-1 * Quads(i, j).NestTrace)
       End If
      End If
      picTerreno(0).CurrentX = ((Foods(b).i - 1) * GridSquare + GridSquare * 14 / 20) * fx
      picTerreno(0).CurrentY = ((Foods(b).j - 1) * GridSquare + GridSquare * 15 / 20) * fy
      picTerreno(0).Print Quads(Foods(b).i, Foods(b).j).Ants
      picTerreno(0).ForeColor = RGB(0, 0, 0)
    End If
  End If
 Next b
 
 'Se dibuja el nido:
 'The nest is drawed:
 picMapa(0).FillColor = RGB(50, 50, 50)
 picMapa(0).Circle (NestX, NestY), NestSize / fx, vbBlack
 picTerreno(0).FillColor = RGB(50, 50, 50)
 picTerreno(0).Circle (NestX * fx, NestY * fy), NestSize, RGB(50, 50, 50)
 'picMapa(0).Circle (NestX, NestY), Sqr(NestFood / PI) / fx, RGB(0, 0, 255)
 If Odor = 1 Then
   picTerreno(0).ForeColor = RGB(255, 255, 0)
   picTerreno(0).CurrentX = ((Int(NestX / GridSquare)) * GridSquare + GridSquare / 20) * fx
   picTerreno(0).CurrentY = ((Int(NestY / GridSquare)) * GridSquare + GridSquare / 20) * fy
   If Quads(Int(NestX / GridSquare) + 1, Int(NestY / GridSquare) + 1).Trace > 0 Then
    picTerreno(0).Print Int(Quads(Int(NestX / GridSquare) + 1, Int(NestY / GridSquare) + 1).Trace)
    'picTerreno(0).Print Int(Quads(i, j).FoodTrace)
   ElseIf Quads(Int(NestX / GridSquare) + 1, Int(NestY / GridSquare) + 1).Trace < 0 Then
     picTerreno(0).Print -Int(Quads(Int(NestX / GridSquare) + 1, Int(NestY / GridSquare) + 1).Trace)
     'picTerreno(0).Print Int(-1 * Quads(i, j).NestTrace)
   End If
   picTerreno(0).CurrentX = ((Int(NestX / GridSquare)) * GridSquare + GridSquare * 14 / 20) * fx
   picTerreno(0).CurrentY = ((Int(NestY / GridSquare)) * GridSquare + GridSquare * 15 / 20) * fy
   picTerreno(0).Print Quads(Int(NestX / GridSquare) + 1, Int(NestY / GridSquare) + 1).Ants
   picTerreno(0).ForeColor = RGB(0, 0, 0)
 End If
   
 For a = 0 To NumAnts - 1
  AntXm = (Ants(a).X0 + Ants(a).X1) / 2
  AntYm = (Ants(a).Y0 + Ants(a).Y1) / 2
  
  'Se dibujan las hormigas en el mapa general:
  'Ants are drawed on general map:
  picMapa(0).FillColor = RGB(150, 100, 70)
  picMapa(0).Line (AntXm - AntSize / (2 * fx), AntYm)-(AntXm + AntSize / (2 * fx), AntYm), RGB(150, 100, 70)
  picMapa(0).Line (AntXm, AntYm - AntSize / (2 * fy))-(AntXm, AntYm + AntSize / (2 * fy)), RGB(150, 100, 70)
  picMapa(0).Circle (AntXm, AntYm), AntSize / (2 * fx), RGB(150, 100, 70)
  'picMapa(0).Print a
  'Se dibuja la comida transportada en el mapa general:
  'Carried food is drawed on general map:
  If (Ants(a).Food > 0) Then
   picMapa(0).FillColor = RGB(16, 152, 16)
   picMapa(0).Circle (Ants(a).X1, Ants(a).Y1), Sqr(Ants(a).Food / PI) / fx, RGB(16, 152, 16)
  End If
  picMapa(0).FillColor = RGB(150, 100, 70)
  'picMapa(0).Line (ants(a).x0, ants(a).y0)-(ants(a).x1, ants(a).y1), RGB(200, 100, 0)
  'picMapa(0).Circle (ants(a).x0, ants(a).y0), AntSize / (4 * fx), RGB(200, 100, 0)
  'picMapa(0).Circle (ants(a).x1, ants(a).y1), AntSize / (5 * fx), RGB(200, 100, 0)
   
  If ((Ants(a).X0 > picTerreno(0).ScaleLeft / fx And Ants(a).X0 < (picTerreno(0).ScaleLeft + picTerreno(0).ScaleWidth) / fx And Ants(a).Y0 > picTerreno(0).ScaleTop / fx And Ants(a).Y0 < (picTerreno(0).ScaleTop + picTerreno(0).ScaleHeight) / fx) _
       Or (Ants(a).X1 > picTerreno(0).ScaleLeft / fx And Ants(a).X1 < (picTerreno(0).ScaleLeft + picTerreno(0).ScaleWidth) / fx And Ants(a).Y1 > picTerreno(0).ScaleTop / fx And Ants(a).Y1 < (picTerreno(0).ScaleTop + picTerreno(0).ScaleHeight) / fx)) Then
  'Se dibuja la comida transportada en el mapa local:
  'Carried food is drawed on local map:
   If (Ants(a).Food > 0) Then
    picTerreno(0).FillColor = RGB(16, 152, 16)
    picTerreno(0).Circle ((Ants(a).X1 + (Ants(a).X1 - Ants(a).X0) / 4) * fx, (Ants(a).Y1 + (Ants(a).Y1 - Ants(a).Y0) / 4) * fy), Sqr(Ants(a).Food / PI), RGB(0, 0, 0)
   End If
   'Se dibujan las hormigas visibles en el mapa local:
   'Ants visible on local map are drawed:
    'abdomen:
    picTerreno(0).FillColor = RGB(160, 100, 70)
    If a = 0 Then
     picTerreno(0).Circle ((Ants(a).X0 + (Ants(a).X1 - Ants(a).X0) / 4) * fx, (Ants(a).Y0 + (Ants(a).Y1 - Ants(a).Y0) / 4) * fy), Ants(a).Size / 4 + Sqr(NestFood / PI), RGB(0, 0, 0)
    Else
     picTerreno(0).Circle ((Ants(a).X0 + (Ants(a).X1 - Ants(a).X0) / 6) * fx, (Ants(a).Y0 + (Ants(a).Y1 - Ants(a).Y0) / 6) * fy), Ants(a).Size / 6, RGB(0, 0, 0)
    End If
    'head and thorax:
    picTerreno(0).FillColor = RGB(160, 100, 70)
    picTerreno(0).Circle ((Ants(a).X1 - (Ants(a).X1 - Ants(a).X0) * (11 / 28)) * fx, (Ants(a).Y1 - (Ants(a).Y1 - Ants(a).Y0) * (11 / 28)) * fy), Ants(a).Size / 7, RGB(0, 0, 0)
    picTerreno(0).FillColor = RGB(160, 100, 70)
    picTerreno(0).Circle (Ants(a).X1 * fx, Ants(a).Y1 * fy), Ants(a).Size / 5, RGB(0, 0, 0)
  End If
  
  'Se calculan variables auxiliares para rotacion y traslacion de hormigas:
  'Auxiliary variables for ants translation and rotation are calculated:
   k = Int(Rnd * (2)) + 2
   AntX00 = Ants(a).X0
   AntY00 = Ants(a).Y0
   AntX10 = Ants(a).X1
   AntY10 = Ants(a).Y1
   
   If a <> 0 Then
   AntXc = AntX00 + (AntX10 - AntX00) * 1 / k
   AntYc = AntY00 + (AntY10 - AntY00) * 1 / k
   Else
    AntXc = NestX
    AntYc = NestY
    Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 60 * (PI / 180)
   End If
   
   'Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 60 * (PI / 180)
    
  'Se Determina posicion actual de la hormiga en la grilla:
  'Current ant position on the grid is calculated:
  Currenti = Int(Ants(a).X1 / GridSquare) + 1
  Currentj = Int(Ants(a).Y1 / GridSquare) + 1
  
  'Rotacion:
  'Rotation:
   If ((AntXc + (AntX10 - AntXc) * Cos(Ants(a).Theta) - (AntY10 - AntYc) * Sin(Ants(a).Theta)) >= 0 _
       And (AntXc + (AntX10 - AntXc) * Cos(Ants(a).Theta) - (AntY10 - AntYc) * Sin(Ants(a).Theta)) <= picMapa(0).Width _
       And (AntYc + (AntX10 - AntXc) * Sin(Ants(a).Theta) + (AntY10 - AntYc) * Cos(Ants(a).Theta)) >= 0 _
       And (AntYc + (AntX10 - AntXc) * Sin(Ants(a).Theta) + (AntY10 - AntYc) * Cos(Ants(a).Theta)) <= picMapa(0).Height) Then
    Ants(a).X0 = AntXc + (AntX00 - AntXc) * Cos(Ants(a).Theta) - (AntY00 - AntYc) * Sin(Ants(a).Theta)
    Ants(a).Y0 = AntYc + (AntX00 - AntXc) * Sin(Ants(a).Theta) + (AntY00 - AntYc) * Cos(Ants(a).Theta)
    Ants(a).X1 = AntXc + (AntX10 - AntXc) * Cos(Ants(a).Theta) - (AntY10 - AntYc) * Sin(Ants(a).Theta)
    Ants(a).Y1 = AntYc + (AntX10 - AntXc) * Sin(Ants(a).Theta) + (AntY10 - AntYc) * Cos(Ants(a).Theta)
   End If
     
  'Traslacion:
  'Translation:
  dx = (Ants(a).X1 - Ants(a).X0) / k
  dy = (Ants(a).Y1 - Ants(a).Y0) / k
  
  Futurei = Int((Ants(a).X1 + dx) / GridSquare) + 1
  Futurej = Int((Ants(a).Y1 + dy) / GridSquare) + 1
  
  If a <> 0 Then
   'Si al trasladarlas caen dentro del area de mapa:
   'If after translation ants remains on map limits:
   If ((Futurei >= 1) And (Futurei <= GridDimX) And (Futurej >= 1) And (Futurej <= GridDimY)) Then
    'Reglas empiricas para el movimiento de las hormigas:
    'Empirical rules for ants movement:
    If (Quads(Futurei, Futurej).Ants <= ((GridSquare * fx) ^ 2 / AntArea) / 2) And _
     (Currenti = Futurei Or Currentj = Futurej) And _
     ( _
     (Currenti = Futurei And Currentj = Futurej) _
     Or (Ants(a).Food = 0 And Quads(Currenti, Currentj).Trace <= 0 And MaxTraceQuads(Currenti, Currentj) <= 0) _
     Or (Ants(a).Food = 0 And Quads(Futurei, Futurej).Trace = MaxTraceQuads(Currenti, Currentj)) _
     Or (Ants(a).Food = 0 And Quads(Currenti, Currentj).Trace = MaxTraceQuads(Currenti, Currentj) And Quads(Futurei, Futurej).Trace = MinTraceQuads(Currenti, Currentj)) _
     Or (Ants(a).Food > 0 And MinTraceQuads(Currenti, Currentj) = 0 And Quads(Futurei, Futurej).Trace > 0) _
     Or (Ants(a).Food > 0 And Quads(Futurei, Futurej).Trace = -frmOptions.sliNestTrace.Value) _
     Or (Ants(a).Food > 0 And (MinTraceQuads(Currenti, Currentj) <> 0 Or Quads(Futurei, Futurej).Trace < 0) And Quads(Futurei, Futurej).Trace = MinTraceQuads(Currenti, Currentj)) _
     Or (Ants(a).Food > 0 And Quads(Currenti, Currentj).Trace = MinTraceQuads(Currenti, Currentj) And Quads(Futurei, Futurej).Trace = MaxTraceQuads(Currenti, Currentj) And Quads(Currenti, Currentj).Trace > -frmOptions.sliNestTrace.Value) _
     ) Then
     'Or (Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace) _
     'Or (Quads(Currenti, Currentj).Trace = 0) _
     'Or (Ants(a).Food > 0 And Quads(Futurei, Futurej).Trace <> 0 And Quads(Futurei, Futurej).Trace = MinTraceQuads(Currenti, Currentj)) _

      Ants(a).X0 = Ants(a).X0 + dx
      Ants(a).Y0 = Ants(a).Y0 + dy
      Ants(a).X1 = Ants(a).X1 + dx
      Ants(a).Y1 = Ants(a).Y1 + dy
      Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 60 * (PI / 180)
      
      'Ants quantity on active quadrants is updated:
      'If ((Currenti <> Futurei) Or (Currentj <> Futurej)) Then
       Quads(Currenti, Currentj).Ants = Quads(Currenti, Currentj).Ants - 1
       Quads(Futurei, Futurej).Ants = Quads(Futurei, Futurej).Ants + 1
      'End If
      
      If Ants(a).Food = 0 Then 'Ants carrying no food
       If Quads(Currenti, Currentj).Trace <= 0 Then 'Odor Trace, or no Trace
        If Quads(Futurei, Futurej).Trace > 0 Then
         If Abs(Quads(Currenti, Currentj).Trace) < Abs(Quads(Futurei, Futurej).Trace) _
           And Quads(Currenti, Currentj).Trace <> -frmOptions.sliNestTrace.Value Then
          Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace - 1
         ElseIf Abs(Quads(Currenti, Currentj).Trace) > Abs(Quads(Futurei, Futurej).Trace) Then
          'Quads(Futurei, Futurej).Trace = Quads(Currenti, Currentj).Trace + 1
         End If
        ElseIf Quads(Futurei, Futurej).Trace <= 0 Then
         If Quads(Futurei, Futurej).Trace < Quads(Currenti, Currentj).Trace Then
          Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace + 1
         ElseIf Quads(Futurei, Futurej).Trace > Quads(Currenti, Currentj).Trace Then
          Quads(Futurei, Futurej).Trace = Quads(Currenti, Currentj).Trace + 1
         End If
        End If
       Else 'Odor trace present on current ant position
        If Quads(Currenti, Currentj).Trace > 0 Then 'Food Trace
         Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 5 * (PI / 180)
         If Quads(Currenti, Currentj).Trace < Quads(Futurei, Futurej).Trace Then
          Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace - 1
         End If
        Else
          If Quads(Currenti, Currentj).Trace > Quads(Futurei, Futurej).Trace Then
           Quads(Futurei, Futurej).Trace = Quads(Currenti, Currentj).Trace
         End If
        End If
       End If
      ElseIf Ants(a).Food > 0 Then 'Ants carrying food
        If Quads(Currenti, Currentj).Trace <= 0 And Quads(Currenti, Currentj).Trace >= Quads(Futurei, Futurej).Trace Then
         If Quads(Futurei, Futurej).Trace <> 0 Then
          Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 5 * (PI / 180)
          If Quads(Currenti, Currentj).Trace <> -frmOptions.sliNestTrace Then
           Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace + 1
          End If
         End If
        Else
         If Quads(Currenti, Currentj).Trace < 0 And Quads(Futurei, Futurej).Trace > 0 Then
          If Abs(Quads(Currenti, Currentj).Trace) < Abs(Quads(Futurei, Futurej).Trace) Then
           Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace - 1
          End If
         Else
          If Quads(Currenti, Currentj).Trace > 0 Then
           Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 5 * (PI / 180)
           If Quads(Futurei, Futurej).Trace < Quads(Currenti, Currentj).Trace Then
              If Abs(Quads(Currenti, Currentj).Trace) > Abs(Quads(Futurei, Futurej).Trace) _
               And Quads(Futurei, Futurej).Trace <> -frmOptions.sliNestTrace.Value Then
               Quads(Futurei, Futurej).Trace = Quads(Currenti, Currentj).Trace - 1
              Else
               'Quads(Currenti, Currentj).Trace = Quads(Futurei, Futurej).Trace + 1
              End If
           End If
          End If
         End If
        End If
      End If
    Else
     Futurei = Int((Ants(a).X1 - dx) / GridSquare) + 1
     Futurej = Int((Ants(a).Y1 - dy) / GridSquare) + 1
     If ((Currenti <> Futurei) Or (Currentj <> Futurej)) Then
      Quads(Currenti, Currentj).Ants = Quads(Currenti, Currentj).Ants - 1
      Quads(Futurei, Futurej).Ants = Quads(Futurei, Futurej).Ants + 1
     End If
     Ants(a).X0 = Ants(a).X0 - dx
     Ants(a).Y0 = Ants(a).Y0 - dy
     Ants(a).X1 = Ants(a).X1 - dx
     Ants(a).Y1 = Ants(a).Y1 - dy
     Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 60 * (PI / 180)
    End If
   Else 'Si no cumplen las condiciones anteriores, "rebotan":
     'If previous conditions are not satisfied, they "bounce":
     Futurei = Int((Ants(a).X1 - dx) / GridSquare) + 1
     Futurej = Int((Ants(a).Y1 - dy) / GridSquare) + 1
     If ((Futurei >= 1) And (Futurei <= GridDimX) And (Futurej >= 1) And (Futurej <= GridDimY)) Then
      If ((Currenti <> Futurei) Or (Currentj <> Futurej)) Then
       Quads(Currenti, Currentj).Ants = Quads(Currenti, Currentj).Ants - 1
       Quads(Futurei, Futurej).Ants = Quads(Futurei, Futurej).Ants + 1
      End If
      Ants(a).X0 = Ants(a).X0 - dx
      Ants(a).Y0 = Ants(a).Y0 - dy
      Ants(a).X1 = Ants(a).X1 - dx
      Ants(a).Y1 = Ants(a).Y1 - dy
      Ants(a).Theta = (-1) ^ (Int(Rnd * 2) + 1) * Rnd * 60 * (PI / 180)
     End If
   End If
  End If
  
  'Se entrega la comida en el nido:
  'Food is delivered into the nest:
  If a = 0 Then
   Quads(Int(Ants(a).X0 / GridSquare) + 1, Int(Ants(a).Y0 / GridSquare) + 1).Trace = -frmOptions.sliNestTrace.Value
  End If
  If (Sqr(((NestX * fx) - Ants(a).X1 * fx) ^ 2 + ((NestY * fy) - Ants(a).Y1 * fy) ^ 2) < NestSize) And (Currenti = Futurei And Currentj = Futurej) Then
   Quads(Currenti, Currentj).Trace = -frmOptions.sliNestTrace.Value
   If Ants(a).Food > 0 Then
    NestFood = NestFood + Ants(a).Food
    Ants(a).Food = 0
   End If
  End If
 Next a
  
 'El programa termina al entregar todo el alimento en el nido:
 'The application ends when all the food has been delivered:
 FoodCarried = 0
 For a = 0 To NumAnts - 1
  FoodCarried = FoodCarried + Ants(a).Food
 Next a
 If FoodQuad = 0 And FoodCarried = 0 Then
  MsgBox "All food sources have been gathered."
  frmAntsSimulator.btnStop = True
 End If
 
 'Se verifica que la cantidad de hormigas no exceda la capacidad maxima:
 'Ants quantity not exceeding maximum capacity is verified:
 If NumAnts = MaxAnt Then
  MsgBox "Maximum ant population has been reached."
  frmAntsSimulator.btnStop = True
 End If
 
 'Se dibuja el area activa en el mapa general:
 'Active area is drawed on the general map:
 picMapa(0).FillStyle = 1
 picMapa(0).Line (picTerreno(0).ScaleLeft / fx, picTerreno(0).ScaleTop / fy)-(picTerreno(0).ScaleLeft / fx + picTerreno(0).ScaleWidth / fx - 1, picTerreno(0).ScaleTop / fy + picTerreno(0).ScaleHeight / fy - 1), RGB(255, 0, 0), B
 'picMapa(0).Line (HScroll1.Value / fx - picMapa(0).ScaleWidth / 6, VScroll1.Value / fx - picMapa(0).ScaleHeight / 6)-(HScroll1.Value / fx + picMapa(0).ScaleWidth / 6, VScroll1.Value / fy + picMapa(0).ScaleHeight / 6), RGB(255, 0, 0), B
 picMapa(0).FillStyle = 0
 
 'Se generan hormigas nuevas:
 'New ants are generated:
   If NestFood >= MaxBite Then
   'If NestFood > PI * (AntSize / 2) ^ 2 / 3 Then
    NumAnts = NumAnts + 1
    Quads(Int(GridDimX / 2) + 1, Int(GridDimY / 2) + 1).Ants = Quads(Int(GridDimX / 2) + 1, Int(GridDimY / 2) + 1).Ants + 1
    Ants(NumAnts - 1).X0 = NestX
    Ants(NumAnts - 1).Y0 = NestY
    Ants(NumAnts - 1).Size = AntSize
    Ants(NumAnts - 1).X1 = Ants(NumAnts - 1).X0 + (-1) ^ (Int(Rnd * 2) + 1) * (Rnd * AntSize / fx)
    Ants(NumAnts - 1).Y1 = Ants(NumAnts - 1).Y0 + (-1) ^ (Int(Rnd * 2) + 1) * Sqr((AntSize / fx) ^ 2 - (Ants(NumAnts - 1).X1 - Ants(NumAnts - 1).X0) ^ 2)
    Ants(NumAnts - 1).Theta = 0
    AntZ(NumAnts - 1) = 0
    Ants(NumAnts - 1).Food = 0
    NestFood = NestFood - MaxBite
    'NestFood = NestFood - PI * (AntSize / 2) ^ 2 / 3
   End If
 lblPop.Caption = NumAnts
 
 'Animacion mas fluida:
 'More fluid animation (without flickering):
 picMapa(1).Visible = False
 picTerreno(1).Visible = False
 picMapa(0).AutoRedraw = True
 picTerreno(0).AutoRedraw = True

End Sub

Private Sub VScroll1_Change()
 picTerreno(0).ScaleTop = VScroll1.Value - 1 / 2 * picTerreno(0).ScaleHeight
 'picTerreno(0).ScaleTop = VScroll1.Value - picMapa(0).ScaleHeight * fy / 6
End Sub
