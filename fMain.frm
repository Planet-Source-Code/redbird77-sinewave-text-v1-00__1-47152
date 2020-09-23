VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SineWave Text"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSWOptions 
      Caption         =   "SineWave Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   5895
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -10
         SelStart        =   4
         Value           =   4
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -8
         Max             =   8
         SelStart        =   2
         Value           =   2
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   45
         SmallChange     =   15
         Min             =   -180
         Max             =   180
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -12
         Max             =   12
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   31
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -10
         SelStart        =   4
         Value           =   4
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   32
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -8
         Max             =   8
         SelStart        =   2
         Value           =   2
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   33
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   45
         SmallChange     =   15
         Min             =   -180
         Max             =   180
         SelStart        =   90
         Value           =   90
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -12
         Max             =   12
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   8
         Left            =   4080
         TabIndex        =   35
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -10
         SelStart        =   4
         Value           =   4
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   36
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -8
         Max             =   8
         SelStart        =   2
         Value           =   2
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   37
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   45
         SmallChange     =   15
         Min             =   -180
         Max             =   180
         SelStart        =   180
         Value           =   180
      End
      Begin MSComctlLib.Slider sldWave 
         Height          =   255
         Index           =   11
         Left            =   4080
         TabIndex        =   38
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         Min             =   -12
         Max             =   12
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(0)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5160
         TabIndex        =   49
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(180°)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5160
         TabIndex        =   48
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   47
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(0)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   45
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(90°)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   43
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   42
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSineWave 
         Alignment       =   2  'Center
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   41
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblSineWave 
         Alignment       =   2  'Center
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblSineWave 
         Alignment       =   2  'Center
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   39
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(0)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(0°)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSldValue 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSldCaption 
         Caption         =   "Shift V"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblSldCaption 
         Caption         =   "Shift H"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblSldCaption 
         Caption         =   "Frequency"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSldCaption 
         Caption         =   "Amplitude"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4560
      Width           =   4095
   End
   Begin VB.PictureBox picSine 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Frame fraOptions 
      Height          =   4815
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtCycles 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset Defaults"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoCopy 
         Caption         =   "Auto Copy to Clipboard"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblRefresh 
         Caption         =   "[R]"
         Height          =   255
         Left            =   1200
         TabIndex        =   50
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblOptsCaption 
         AutoSize        =   -1  'True
         Caption         =   "Cycles"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblFuture 
         Caption         =   "Presets will go here.  Not implemented yet.  Also triangular waves will come soon to produce better yellows."
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblOptions 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   11
         Top             =   0
         Width           =   720
      End
   End
   Begin RichTextLib.RichTextBox rtfPreview 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393217
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"fMain.frx":000C
   End
   Begin RichTextLib.RichTextBox rtfInput 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"fMain.frx":00DA
   End
   Begin RichTextLib.RichTextBox rtfOutput 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"fMain.frx":01A8
   End
   Begin VB.Label lblCaption 
      Caption         =   "SineWave Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblCaption 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblCaption 
      Caption         =   "HTML Output"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblCaption 
      Caption         =   "Text Input"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblCoords 
      Caption         =   "Coords"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      Caption         =   "http://home.earthlink.net/~redbird77"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   7200
      Width           =   3255
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' file    : fMain.frm
' revised : 03 July 24
' project : pSineWaveText.vbp
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77
' about   : see the README.txt file

' ##TODO## Organize into classes.
' Create IWave interface, and CSineWave, CTriangleWave, & CSawToothWave.
' Create CSineWaveText that has a .SineWave() collection contaning 3 SineWave
' classes.

Option Explicit

Private Type tPhaseShift
    Horizontal As Double
    Vertical   As Double
End Type

Private Enum eColor
    None = -1
    Red = 0
    Green = 1
    Blue = 2
End Enum

Private Type tSineWave
    Amplitude  As Double
    Frequency  As Double
    PhaseShift As tPhaseShift
    Color      As eColor
    Data(359)  As Double
End Type

Private SineWave(2) As tSineWave

Private Sub Form_Load()
   
Dim i As Integer

    rtfInput.Text = """What's so funny about peace, love, and " & _
                    "understanding?"" - Nick Lowe"

    'Set up default SineWaves properties.
    For i = 0 To 2
        With SineWave(i)
            .Amplitude = 1
            .Frequency = 1
            .PhaseShift.Horizontal = (TPI / 4) * i
            .PhaseShift.Vertical = 0
            .Color = i
        End With
    Next
    
    ' Draw all the default SineWaves.
    SineWavesCollection_Draw
    SineWaveText_Draw
    
    With lblURL
        .ForeColor = &HC00000
        .MousePointer = 99
        .MouseIcon = LoadResPicture(101, vbResCursor)
    End With
    
End Sub

Private Sub SineWavesCollection_Draw()

' Draws all three SineWaves and the corresponing gradient at the same time.

Dim i         As Integer
Dim iDeg      As Integer
Dim dCurY     As Double
Dim dOldX     As Double
Dim dOldY     As Double
Dim dColor(2) As Double

    ' Set the scale
    picSine.Scale (0, 1)-(360, -1)
    picColor.Scale (0, 1)-(360, -1)
    
    ' For each SinWave...
    For i = 0 To 2
        ' For each degree in a SineWave...
        For iDeg = 0 To 359

            With SineWave(i)
                ' Calculate the Y-value.
                ' It is a simple y = sin(x) function.
                
                ' Search http://mathforum.org/dr.math/ for good beginner to
                ' advanced help on trig and other aspects of math.
                
                dCurY = .Amplitude * _
                        Sin(DegToRad(iDeg) * .Frequency + .PhaseShift.Horizontal) + _
                        .PhaseShift.Vertical
                        
                ' Store the Y-value at the current degree for further processing.
                .Data(iDeg) = dCurY
                
                ' Draw segment of SineWave.
                If iDeg Then picSine.Line (dOldX, dOldY)-(iDeg, dCurY), _
                                           255 * (2 ^ (.Color * 8))
            
                ' Store old coordinates.
                dOldX = iDeg: dOldY = dCurY
            End With
        Next
    Next
    
    ' Now draw the color that is represented by the values of the 3
    ' SineWaves at any given degree.
    For iDeg = 0 To 359
        For i = 0 To 2
            ' Normalize the color componet value so it's in the [0,255] range.
            dColor(i) = Normalize(SineWave(i).Data(iDeg))
        Next
        
        ' Draw segment of gradient.
        picColor.Line (iDeg, 1)-Step(0, picColor.ScaleHeight), _
                       RGB(dColor(0), dColor(1), dColor(2))
    Next

' ##TODO - Use do while loop and fix last bit not printing.
' ##TODO - Decide whether to store the un/normalized data in the array.
'          Or create a seperate property like .NormalizedData().

End Sub

Private Sub SineWaveText_Draw()

' Here the color is determined by combining the values of the three
' SineWaves at a given point.  Each character's color is determined by
' it's position in the string.

Dim i         As Integer
Dim j         As Integer
Dim dCycles   As Double
Dim dRad      As Double
Dim bytChr()  As Byte
Dim sChr      As String
Dim sChrHtm   As String
Dim dColor(2) As Double
Dim lColor    As Long
Dim fsHTM     As CFastString

    If Len(rtfInput.Text) < 2 Then
        'MsgBox "You must enter at least two characters.", vbExclamation
        Exit Sub
    End If
    
    Set fsHTM = New CFastString
    rtfPreview.Text = "": rtfOutput.Text = ""
    dCycles = Val(txtCycles.Text)
    If dCycles = 0 Then dCycles = 1
    
    bytChr = StrConv(rtfInput.Text, vbFromUnicode)

    For i = 0 To UBound(bytChr)
        sChr = Chr$(bytChr(i))
        
        If Not IsPrint(bytChr(i)) Then
            fsHTM.Append sChr
        Else
            ' This could also be calculated by running through the
            ' SineWave(i).Data() array.  This way you would not have to
            ' re-calculate the sine.

        
            ' ##TODO## Deal with multiple spaces and linebreaks.
            ' dRad = (iPrintableIndex / iPrintableCount) * (TPI * dCycles)
            dRad = (i / UBound(bytChr)) * (TPI * dCycles)
    
            For j = 0 To 2
                dColor(j) = ( _
                            SineWave(j).Amplitude * _
                            Sin(dRad * SineWave(j).Frequency + _
                            SineWave(j).PhaseShift.Horizontal) + _
                            SineWave(j).PhaseShift.Vertical _
                            )
                            
                dColor(j) = Normalize(dColor(j))
            Next
            
            lColor = RGB(dColor(0), dColor(1), dColor(2))
            
            sChrHtm = sChr
            
            If sChrHtm = ">" Then
                sChrHtm = "&gt;"
            ElseIf sChrHtm = "<" Then
                sChrHtm = "&lt;"
            ElseIf sChrHtm = "&" Then
                sChrHtm = "&amp;"
            End If
            
            ' Create the HTML font tag.
            fsHTM.Append "<font color=""" & LNGtoHEX(lColor) & """>" & _
                         sChrHtm & "</font>"
        End If

        With rtfPreview
            .SelStart = Len(.Text)
            .SelColor = lColor
            .SelText = sChr
        End With
    Next
        
    rtfOutput.Text = fsHTM.Buffer
    
    Set fsHTM = Nothing
    
    If chkAutoCopy.Value = vbChecked Then
        Clipboard.Clear
        Clipboard.SetText rtfOutput.Text
    End If
    
End Sub

Private Sub cmdMakeText_Click()
    SineWaveText_Draw
End Sub

Private Sub cmdReset_Click()

Dim i As Integer

    For i = 0 To sldWave.UBound
        Select Case i
            Case 0, 4, 8: sldWave(i).Value = 4  ' Amplitude
            Case 1, 5, 9: sldWave(i).Value = 2  ' Frequency
            Case 2, 6, 10: sldWave(i).Value = RadToDeg((TPI / 4) * (i \ 4)) ' PhaseShift.Horizontal
            Case 3, 7, 11: sldWave(i).Value = 0 ' PhaseShift.Vertical
        End Select
    Next
    
    txtCycles.Text = "1"
    rtfOutput.Text = "": rtfInput.Text = "": rtfPreview.Text = ""
    rtfInput.SetFocus
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblRefresh_Click()
    SineWaveText_Draw
End Sub

Private Sub rtfInput_Change()
    SineWaveText_Draw
End Sub

Private Sub sldWave_Scroll(Index As Integer)

' See comments in the sld_Change event procedure.

Dim i            As Integer ' base
Dim sTmpPhase    As String
Dim eActiveColor As eColor

    eActiveColor = (Index \ 4)

    ' Change properties of chosen SineWave (r, g, or b), then render.
    With SineWave(eActiveColor)
    
    ' -- Amplitude --------------------------------------------------
        i = eActiveColor * 4
        .Amplitude = CDbl(sldWave(i).Value / 4)
        lblSldValue(i).Caption = "(" & CStr(.Amplitude) & ")"
        sldWave(i).Text = .Amplitude

    ' -- Frequency --------------------------------------------------
        i = i + 1
        .Frequency = CDbl(sldWave(i).Value / 2)
        lblSldValue(i).Caption = "(" & CStr(.Frequency) & ")"
        sldWave(i).Text = .Frequency

    ' -- PhaseShift --------------------------------------------------
        i = i + 1
        .PhaseShift.Horizontal = DegToRad(sldWave(i).Value)
        sTmpPhase = CStr(sldWave(i).Value) & Chr$(176)
        lblSldValue(i).Caption = "(" & sTmpPhase & ")"
        sldWave(i).Text = sTmpPhase
        
        i = i + 1
        .PhaseShift.Vertical = CDbl(sldWave(i).Value / 6)
        sTmpPhase = Format$(.PhaseShift.Vertical, "0.00")
        lblSldValue(i).Caption = "(" & sTmpPhase & ")"
        sldWave(i).Text = sTmpPhase
        
    End With

End Sub

Private Sub sldWave_Change(Index As Integer)
' When user changes any wave property, change the property of the module-
' level SineWave.  Next re-render all SineWaves.  The same process happens
' in the sldWave_Scroll event, but the SineWaves are not re-rendered there,
' since the Scroll event also causes a Change event.

Dim i            As Integer
Dim sTmpPhase    As String
Dim eActiveColor As eColor

    eActiveColor = (Index \ 4)

    ' Change properties of chosen SineWave (r, g, or b), then render.
    With SineWave(eActiveColor)
    
    ' -- Amplitude --------------------------------------------------
        i = eActiveColor * 4
        .Amplitude = CDbl(sldWave(i).Value / 4)
        lblSldValue(i).Caption = "(" & CStr(.Amplitude) & ")"
        sldWave(i).Text = .Amplitude

    ' -- Frequency --------------------------------------------------
        i = i + 1
        .Frequency = CDbl(sldWave(i).Value / 2)
        lblSldValue(i).Caption = "(" & CStr(.Frequency) & ")"
        sldWave(i).Text = .Frequency

    ' -- PhaseShift --------------------------------------------------
        i = i + 1
        .PhaseShift.Horizontal = DegToRad(sldWave(i).Value)
        sTmpPhase = CStr(sldWave(i).Value) & Chr$(176)
        lblSldValue(i).Caption = "(" & sTmpPhase & ")"
        sldWave(i).Text = sTmpPhase

        i = i + 1
        .PhaseShift.Vertical = CDbl(sldWave(i).Value / 6)
        sTmpPhase = Format$(.PhaseShift.Vertical, "0.00")
        lblSldValue(i).Caption = "(" & sTmpPhase & ")"
        sldWave(i).Text = sTmpPhase

    End With

    ' Re-render all SineWaves.  Could use a buffering technique or memory
    ' device context here to reduce flicker.
    picSine.Cls
    SineWavesCollection_Draw
    SineWaveText_Draw
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fMain = Nothing
End Sub

Private Sub picSine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCoords.Caption = "X: " & X & vbCrLf & "Y: " & Y
End Sub

'--------------------------------------------------------------------
' -- Hyperlink Functions --------------------------------------------
'--------------------------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With lblURL
        .Font.Underline = False
        .ForeColor = &HC00000
    End With

End Sub

Private Sub lblURL_Click()

    Dim lRet As Long

    lRet = ShellExecute(Me.hwnd, "open", lblURL.Caption, _
                        vbNullString, vbNullString, SW_SHOWNORMAL)

    If lRet <= 32 Then
        MsgBox "Errors occured.  Could not open link.", _
                vbExclamation, "Error"
    End If

End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With lblURL
        ' When mouse is over link, change link color, underline, and
        ' change cursor to hand.
        .Font.Underline = True
        .ForeColor = &H66FF&
    End With

End Sub

'--------------------------------------------------------------------
' -- Misc Functions (DO WORK, BUT NOT IMPLEMENTED YET) --------------
'--------------------------------------------------------------------
Private Sub SineWave_Draw(ByRef SineWave As tSineWave)

' Draw individual SineWave.

Dim dY As Double

' -- Degrees ----------------------
Dim iDeg As Integer
Dim oldX As Double, oldY As Double

    picSine.Scale (0, 1)-(360, -1)
    
    For iDeg = 0 To 359
    
        With SineWave

            dY = .Amplitude * _
                 Sin(DegToRad(iDeg) * .Frequency + .PhaseShift.Horizontal) + _
                 .PhaseShift.Vertical
            
            .Data(iDeg) = dY
            
            If iDeg Then picSine.Line (oldX, oldY)-(iDeg, dY), 255 * (2 ^ (.Color * 8))
            
            oldX = iDeg: oldY = dY
            
        End With
    Next
    
' -- Radians ----------------------
' Dim dRad As Double
'
'    picSine.Scale (0, 1)-(TPI, -1)
'
'    For dRad = 0 To TPI Step 0.001
'        With SineWave
'
'            dY = .Amplitude * _
'                 Sin(dRad * .Frequency + .PhaseShift.Horizontal) + _
'                 .PhaseShift.Vertical
'
'            picSine.PSet (dRad, dY), 255 * (2 ^ (.Color * 8))
'
'        End With
'    Next

End Sub

Private Sub cmdTriangleWave_Click()
    TriangleWave_Draw
End Sub

Private Sub TriangleWave_Draw()

Dim dRad As Double
Dim dDisplacement As Double
Dim dAmplitude As Double: dAmplitude = 1
Dim dFrequency As Double: dFrequency = 2
Dim i As Integer, iNumWaves As Integer: iNumWaves = 100

    picSine.Scale (0, 2)-(TPI, -2)

    For dRad = 0 To TPI Step 0.001
    
        dDisplacement = 0
        
        ' ex: NumWaves = 7
        For i = 1 To (iNumWaves * 2) Step 2
            ' i   = 1,3,5,7,9,11,13
            ' i^2 = 1,9,25,49,81,121,169
            ' Fourier series.
            dDisplacement = dDisplacement + _
                            (((1 / (i ^ 2)) * Cos(i * dFrequency * dRad)) * dAmplitude)
        Next
        
        picSine.PSet (dRad, dDisplacement), vbRed
        
    Next

End Sub
