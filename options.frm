VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Slider sliFoodRate 
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   20
      SmallChange     =   10
      Max             =   100
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton btnAcept 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox NumAnts 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox AntSize 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox GridDimY 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox GridDimX 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin MSComctlLib.Slider sliFoodTrace 
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   3000
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   10
      Min             =   1
      Max             =   500
      SelStart        =   1
      TickFrequency   =   50
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sliNestTrace 
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   4080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   99
      SmallChange     =   10
      Min             =   1
      Max             =   500
      SelStart        =   1
      TickFrequency   =   50
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Workers Lenght:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "mm"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Nest Scent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Weak"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Strong"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Food Scent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Weak"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Strong"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Grid Dimentions    :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Abundant"
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Scarce"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Food Quantity"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Population:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnAcept_Click()
 frmOptions.GridDimY = frmOptions.GridDimX
 frmOptions.Hide
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub



Private Sub GridDimX_Change()
 GridDimY.Text = GridDimX.Text
End Sub


