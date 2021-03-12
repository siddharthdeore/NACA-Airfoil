VERSION 5.00
Begin VB.Form frmCanvas 
   Caption         =   "Foil Maker"
   ClientHeight    =   8844
   ClientLeft      =   1848
   ClientTop       =   480
   ClientWidth     =   16764
   FillStyle       =   3  'Vertical Line
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8844
   ScaleWidth      =   16764
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTool 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1164
      ScaleWidth      =   16644
      TabIndex        =   1
      Top             =   7080
      Width           =   16692
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Chord"
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   37
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtLength 
         Height          =   288
         Left            =   11160
         TabIndex        =   33
         Text            =   "1"
         Top             =   840
         Width           =   372
      End
      Begin VB.OptionButton optType 
         Caption         =   "Symetric"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show XY Marks"
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   32
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.HScrollBar hsAOF 
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   1812
      End
      Begin VB.HScrollBar hsThickness 
         Height          =   252
         LargeChange     =   10
         Left            =   120
         Max             =   60
         Min             =   1
         TabIndex        =   30
         Top             =   360
         Value           =   1
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "NACA 4 Digit Airfoil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11040
         TabIndex        =   2
         Top             =   120
         Width           =   1935
         Begin VB.TextBox txtNACA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   6
            Text            =   "2"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtNACA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   5
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtNACA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   4
            Text            =   "4"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtNACA 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Fill Lines"
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Border"
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Grid"
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtPosition 
         Height          =   375
         Left            =   14280
         TabIndex        =   17
         Text            =   "0.40"
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtCamber 
         Height          =   375
         Left            =   14160
         TabIndex        =   16
         Text            =   "0.1"
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtXX 
         Height          =   375
         Left            =   14400
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtY 
         Height          =   375
         Left            =   14040
         TabIndex        =   14
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtZoom 
         Height          =   375
         Left            =   14040
         TabIndex        =   13
         Text            =   "11000"
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtX 
         Height          =   375
         Left            =   14040
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.HScrollBar hsZoom 
         Height          =   252
         LargeChange     =   10
         Left            =   2040
         Max             =   20000
         Min             =   100
         TabIndex        =   11
         Top             =   360
         Value           =   100
         Width           =   1815
      End
      Begin VB.TextBox txtPoints 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3960
         TabIndex        =   10
         Text            =   "60"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtThick 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.24"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optType 
         Caption         =   "Camberd"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Axis"
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   28
         Top             =   480
         Value           =   1  'Checked
         Width           =   1092
      End
      Begin VB.Label lblAOF 
         Caption         =   "Angle of Attack 0"
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label3 
         Caption         =   "Length"
         Height          =   252
         Left            =   10560
         TabIndex        =   35
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Units"
         Height          =   252
         Left            =   11520
         TabIndex        =   34
         Top             =   840
         Width           =   612
      End
      Begin VB.Label lblStatus 
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   13320
         TabIndex        =   27
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label lblStatus 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   13320
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Camber"
         Height          =   255
         Left            =   13920
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Camber Position"
         Height          =   255
         Left            =   15240
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPts 
         Caption         =   "No of points"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblZoom 
         Caption         =   "Zoom"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblThickness 
         Caption         =   "Thicknes"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Canvas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   15  'Merge Pen Not
      FillStyle       =   0  'Solid
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7332
      ScaleWidth      =   15612
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "siddharthdeore@gmail.com - Siroi.com"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   15495
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolbox 
         Caption         =   "Tool Box"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuShowPoints 
         Caption         =   "Show Points"
      End
      Begin VB.Menu mbuReset 
         Caption         =   "Reset"
         Shortcut        =   {F5}
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuShowBorder 
         Caption         =   "Show Border"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFillLines 
         Caption         =   "Fill Lines"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuShowGrid 
         Caption         =   "Show Grid"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuAxis 
         Caption         =   "Show Axis"
         Checked         =   -1  'True
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuXY 
         Caption         =   "Show XY Marks"
         Checked         =   -1  'True
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Foil Maker"
      End
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public NACA As String
Dim CenterX, CenterY, Points, Zoom As Double
'Dim StartX As Integer, StartY As Integer
'Dim xX, yY, Ymc1, sYl, sYu, sX, cYl, cYu As Variant
'Dim NACA As Variant
Dim t, X, Y, i As Variant
Dim j, k, l, h As Double
Dim Y_Offset, X_Offset, dy, dx As Double
Dim m, p As Variant
Dim pData, pData2 As Variant
Dim ShowBorder, FillLines, ShowGrid, ShowAxis As Boolean
Const pi = 3.14159265358979
Dim thetaD As Double
Dim the2 As Double
Dim base, hyp, perp As Double
Dim isRender As Boolean

Private Sub Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    isRender = False
    Canvas.MousePointer = MousePointerConstants.vbSizePointer
    txtX.Text = X
    txtY.Text = Y
    Else
    Canvas.MousePointer = MousePointerConstants.vbCrosshair
End If
lblStatus(1).Caption = "Current X= " + Str(X) + "Y = " + Str(Y)
NACA = Trim((txtCamber.Text * 10)) + Trim(Int(txtPosition.Text * 10)) + Format(txtThick.Text * 100, "0#")


lblStatus(0).Caption = "X = " + X_Offset + " " + "Y = " + Y_Offset + vbCrLf + "Aerofoil : NACA" + NACA

End Sub


Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
isRender = True
plot
End Sub

Private Sub chkShow_Click(Index As Integer)
plot
End Sub

Public Function Reset_geometry()
txtX.Text = (Canvas.ScaleWidth - 10000) / 2
txtY.Text = Canvas.Height / 2 - picTool.Height
txtThick.Text = 0.12
txtPoints.Text = 60
txtZoom.Text = 10000
lblZoom.Caption = "Zoom = 100%"
hsThickness.Value = 12
hsZoom.Value = 10000
End Function

Private Sub Form_Activate()
Canvas.Width = frmCanvas.ScaleWidth
Canvas.Height = frmCanvas.ScaleHeight
Form_Resize
'hsAOF.Max = pi * 1000
hsAOF.Min = -0.331612557878923 * 1000
hsAOF.Max = 0.331612557878923 * 1000

Reset_geometry
isRender = True
plot
End Sub

Private Sub Form_GotFocus()
plot
End Sub

Private Sub Form_Initialize()
    CenterX = Canvas.Width / 2
    CenterY = Canvas.Height / 2
    Reset_geometry
    optType_Click (0)
    txtNACA(0) = 0
    txtNACA(1) = 0
    plot
End Sub



Public Function plot()
NACA = Trim((txtCamber.Text * 10)) + Trim(Int(txtPosition.Text * 10)) + Format(txtThick.Text * 100, "0#")

Dim length As Integer
length = txtLength.Text 'length of chortd for print purpose

If isRender Then Points = txtPoints.Text Else Points = 18 'for fast render
'Me.Caption = "Ploting Geometry"
Canvas.Refresh
thetaD = hsAOF.Value / 1000
'pData = "   X      Y       Mean Camber Y" + vbCrLf
pData = ""
pData2 = ""
'frmPoints.txtPoints = ""
'Points = Val(txtPoints.Text)

ShowBorder = chkShow(0).Value
FillLines = chkShow(1).Value
ShowGrid = chkShow(2).Value
ShowAxis = chkShow(3).Value
ShowXY = chkShow(4).Value
ShowChord = chkShow(5).Value
Zoom = txtZoom.Text
spacing = Zoom / 6

' Plot XY axis
Canvas.DrawWidth = 1
If ShowXY Then
Canvas.ForeColor = &H40AA00
Canvas.Line (500, Canvas.Height - 4500)-(500, Canvas.Height - 3000)

Canvas.ForeColor = &H1000FF
Canvas.Line (500, Canvas.Height - 3000)-(2000, Canvas.Height - 3000)
End If



If Zoom < 3000 Then spacing = Zoom / 3 ' for controling no of grid line when zoomin and zoomout
If Zoom < 1500 Then spacing = Zoom * 2

X_Offset = txtX.Text
Y_Offset = txtY.Text
t = txtThick.Text
m = txtCamber.Text / 10 ' camber %of chord
p = txtPosition.Text ' camber position %of chord

j = 0
h = 0
k = 0
l = 0

Canvas.ForeColor = vbWhite
Canvas.DrawWidth = 1


PrevX = X_Offset ' for fill line
PrevXl = X_Offset ' for fill line
PrevYu = Y_Offset ' for fill line
PrevYl = Y_Offset ' for fill line

Canvas.ForeColor = &H222222

'Print Grids
If ShowGrid Then
While sarak < Canvas.ScaleWidth * 2 And sarak < Canvas.ScaleHeight * 2
    Canvas.Line (X_Offset + sarak, 0)-(X_Offset + sarak, Canvas.Height)
    Canvas.Line (0, Y_Offset + sarak)-(Canvas.Width, Y_Offset + sarak)
    
    
    Canvas.Line (X_Offset - sarak, 0)-(X_Offset - sarak, Canvas.Height)
    Canvas.Line (0, Y_Offset - sarak)-(Canvas.Width, Y_Offset - sarak)
    
    sarak = sarak + spacing
Wend
End If

'Main Aerofoil Geometry Loop
For i = 0 To (Points - 1)
    xX = (xX + (1 / Points))
    yY = (t / 0.2) * ((0.2969 * Sqr(xX)) - (0.126 * xX) - (0.3516 * xX ^ 2) + (0.2843 * xX ^ 3) - (0.1015 * xX ^ 4))
    If p = 0 Then p = 0.001
    If xX <= p Then
        Ymc1 = (m / (p * p)) * ((2 * p * xX) - (xX * xX))
    Else
        Ymc1 = (m / ((1 - p) ^ 2)) * ((1 - 2 * p) + (2 * p * xX) - (xX * xX))
    End If
    
    'Saves The points to Variable pData for future use
'    pData = pData + FormatNumber(xX * length, 3) + vbTab + FormatNumber(yY * length, 6) + vbTab + FormatNumber(Ymc1 * length, 6) + vbCrLf
    
    pData = FormatNumber(xX * length, 4) + "     " + FormatNumber(yY * length, 4) + vbCrLf + pData
    pData2 = pData2 + FormatNumber(xX * length, 4) + "     " + FormatNumber(-yY * length, 4) + vbCrLf
   
    dist = Abs(Y_Offset - (yY * Zoom))
    
    sX = (xX * Zoom) + X_Offset
    sYl = (yY * Zoom) + Y_Offset
    sYu = (-yY * Zoom) + Y_Offset
    
    base = (sX - X_Offset)
    perp = (yY * Zoom)
    hyp = (base ^ 2 + perp ^ 2) ^ 0.5
'    the2 = Atn(perp / base)
    Ymc1 = Ymc1 * Zoom
    
    sX = (xX * Zoom) * Cos(thetaD) + (yY * Zoom) * Tan(thetaD) + X_Offset
    sXl = (xX * Zoom) * Cos(thetaD) + (-yY * Zoom) * Tan(thetaD) + X_Offset
    sYl = (xX * Zoom) * Sin(thetaD) + Y_Offset + (yY * Zoom)
    sYu = (xX * Zoom) * Sin(thetaD) + Y_Offset + (-yY * Zoom)
    
'    plot symetric
    If optType(0).Value = True Then
    
    If FillLines Then
        'Fill With lines
        Canvas.ForeColor = &H133513
        Canvas.Line (sX, sYu)-(sXl, sYl)
'       Canvas.Line (sX, Y_Offset)-(sX, sYu)
'       Canvas.Line (sX, Y_Offset)-(sX, sYl)
    End If
    
    If ShowBorder Then
        'Border lines
        Canvas.ForeColor = &H506000
        Canvas.Line (PrevX, PrevYu)-(sX, sYu)
        Canvas.Line (PrevXl, PrevYl)-(sXl, sYl)
        
        PrevYl = sYl
        PrevYu = sYu
        PrevX = sX
        PrevXl = sXl
    End If
    
    Canvas.ForeColor = vbWhite
    
    'plot Symetric X and y upp lower
    Canvas.DrawWidth = 2
    Canvas.PSet (sXl, sYl)
    'Canvas.PSet ((xX * Zoom) * Cos(thetaD) + X_Offset, (xX * Zoom) * Sin(thetaD) + Y_Offset)
    Canvas.PSet (sX, sYu)

'Canvas.DrawWidth = 1

' else draw camber profile
    Else
    
        Canvas.ForeColor = vbRed
        cYl = (xX * Zoom) * Sin(thetaD) + (Y_Offset - Ymc1) + (yY * Zoom)
        cYu = (xX * Zoom) * Sin(thetaD) + (Y_Offset - Ymc1) - (yY * Zoom)
            
        If FillLines Then
            'Fill With lines
            Canvas.ForeColor = &H133513
            Canvas.Line (sXl, cYl)-(sX, cYu)
        End If
    
        If ShowBorder Then
        'Border lines

            Canvas.ForeColor = &H506000
            Canvas.Line (PrevX, PrevYu)-(sX, cYu)
            Canvas.Line (PrevXl, PrevYl)-(sXl, cYl)
            
            PrevYl = cYl
            PrevYu = cYu
            PrevX = sX
            PrevXl = sXl
        End If
        
        'plots Camber line
        Canvas.ForeColor = &H989898
        Canvas.PSet (sX, (xX * Zoom) * Sin(thetaD) + (Y_Offset - Ymc1))
    
        Canvas.ForeColor = vbWhite
        Canvas.PSet (sXl, cYl)
        Canvas.PSet (sX, cYu)
    End If
Next i

' Plot Axis

Canvas.DrawWidth = 1
If ShowAxis Then
Canvas.ForeColor = &H404040
Canvas.Line (0, Y_Offset)-(Canvas.Width, Y_Offset)
Canvas.Line (X_Offset, 0)-(X_Offset, Canvas.Height)
End If


' Plot Chordline
If ShowChord Then
Canvas.ForeColor = &H405040
Canvas.Line (X_Offset, Y_Offset)-(X_Offset + Val(Zoom) * Cos(thetaD), (xX * Zoom) * Sin(thetaD) + (Y_Offset - Ymc1))
End If
'Canvas.Circle (X_Offset + Zoom / 2, Y_Offset), Zoom / 2
Canvas.ForeColor = vbWhite

Canvas.PSet (X_Offset, Y_Offset)
Canvas.PSet (Val(X_Offset) + Val(Zoom) * Cos(thetaD) + 10, (xX * Zoom) * Sin(thetaD) + (Y_Offset - Ymc1))

frmPoints.txtPoints.Text = "NACA" + NACA + vbCrLf + Trim(pData) + Trim(pData2)

'Me.Caption = "Airfoil Maker"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "Please feel free to contact siddharthdeore@gmail.com for Suggetions and query....", vbInformation
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
Canvas.Width = frmCanvas.ScaleWidth
Canvas.Height = frmCanvas.ScaleHeight + 1460
picTool.Top = frmCanvas.ScaleHeight - picTool.Height
picTool.Width = frmCanvas.ScaleWidth - 25
lblTitle.Width = frmCanvas.ScaleWidth
plot
End Sub

Private Sub hsAOF_Change()
plot
End Sub

Private Sub hsAOF_Scroll()
lblAOF.Caption = "Angle of Attack " + Str(Round((thetaD / 0.01745329), 1)) + "°"
isRender = False
plot
isRender = True
End Sub

Private Sub hsThickness_Scroll()
txtThick.Text = hsThickness.Value / 100
End Sub

Private Sub hsZoom_Change()
txtZoom.Text = hsZoom.Value
End Sub

Private Sub hsZoom_Scroll()
txtZoom.Text = hsZoom.Value
End Sub
Private Sub mbuReset_Click()
Reset_geometry
plot
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAxis_Click()
    If mnuAxis.Checked = True Then
        mnuAxis.Checked = False
        chkShow(3).Value = 0
    Else
        mnuAxis.Checked = True
        chkShow(3).Value = 1
    End If
End Sub

Private Sub mnuContents_Click()
Dim objWeb As Object
Set objWeb = CreateObject("InternetExplorer.Application")
objWeb.Visible = True
objWeb.navigate CStr(App.Path + "\Help\Help.html"), Null, Null, Null
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFillLines_Click()
    If mnuFillLines.Checked = True Then
        mnuFillLines.Checked = False
        chkShow(1).Value = 0
    Else
        mnuFillLines.Checked = True
        chkShow(1).Value = 1
    End If
End Sub

Private Sub mnuSave_Click()
SavePicture Canvas.Image, App.Path + "\" + Trim(Str(Rnd)) + ".jpg"
End Sub

Private Sub mnuShowBorder_Click()
    If mnuShowBorder.Checked = True Then
        mnuShowBorder.Checked = False
        chkShow(0).Value = 0
    Else
        mnuShowBorder.Checked = True
        chkShow(0).Value = 1
    End If
End Sub

Private Sub mnuShowGrid_Click()
    If mnuShowGrid.Checked = True Then
        mnuShowGrid.Checked = False
        chkShow(2).Value = 0
    Else
        mnuShowGrid.Checked = True
        chkShow(2).Value = 1
    End If
End Sub

Private Sub mnuShowPoints_Click()
frmPoints.Show 1
End Sub

Private Sub mnuToolbox_Click()
    If picTool.Visible = True Then
        picTool.Visible = False
        mnuToolbox.Checked = False
    Else
        picTool.Visible = True
        mnuToolbox.Checked = True
    End If
End Sub

Private Sub mnuXY_Click()
    If mnuXY.Checked = True Then
        mnuXY.Checked = False
        chkShow(4).Value = 0
    Else
        mnuXY.Checked = True
        chkShow(4).Value = 1
    End If
End Sub

Private Sub optType_Click(Index As Integer)
If optType(0).Value = True Then
Frame1.Visible = False
txtCamber.Visible = False
txtPosition.Visible = False
Label1.Visible = False
Label2.Visible = False
'txtNACA(0).Text = 0
'txtNACA(1).Text = 0
Else
Frame1.Visible = True
'txtCamber.Visible = True
'txtPosition.Visible = True
'Label1.Visible = True
'Label2.Visible = True

'txtNACA(0).Text = 1
'txtNACA(1).Text = 4
End If
plot
End Sub






Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KeyCodeConstants.vbKeyDown Then
    txtLength.Text = Val(txtLength.Text - 1)
End If
    If KeyCode = KeyCodeConstants.vbKeyUp Then
    txtLength.Text = Val(txtLength.Text + 1)
End If

End Sub

Private Sub txtNACA_Change(Index As Integer)
On Error Resume Next
If txtNACA(Index).Text > 9 Then txtNACA(Index) = 9
If txtNACA(Index).Text < 1 Then txtNACA(Index) = 0
txtCamber.Text = txtNACA(0).Text / 10
txtPosition.Text = txtNACA(1).Text / 10
txtThick.Text = (txtNACA(2).Text + txtNACA(3)) / 100
plot
End Sub

Private Sub txtNACA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = KeyCodeConstants.vbKeyDown Then
txtNACA(Index).Text = Val(txtNACA(Index).Text - 1)
End If
If KeyCode = KeyCodeConstants.vbKeyUp Then
txtNACA(Index).Text = Val(txtNACA(Index).Text + 1)
End If
End Sub

Private Sub txtPoints_Change()
On Error Resume Next
If txtPoints < 10 Then txtPoints.Text = 10
If txtPoints.Text > 1000 Then txtPoints.Text = 1000
plot
End Sub

Private Sub txtPoints_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KeyCodeConstants.vbKeyUp Then
txtPoints.Text = txtPoints.Text + 1
End If
If KeyCode = KeyCodeConstants.vbKeyDown Then
txtPoints.Text = txtPoints.Text - 1
End If
End Sub


Private Sub txtThick_Change()
lblThickness.Caption = "Thicknes = " + txtThick.Text
plot
End Sub

'Private Sub txtPosition_KeyDown(KeyCode As Integer, Shift As Integer)
'If txtPosition.Text < 0.03 Then txtPosition.Text = 0.03
'If txtPosition.Text > 0.98 Then txtPosition.Text = 0.98
'If KeyCode = KeyCodeConstants.vbKeyUp Then
'    txtPosition.Text = Val(txtPosition.Text) + 0.01
'End If
'
'If KeyCode = KeyCodeConstants.vbKeyDown Then
'    txtPosition.Text = Val(txtPosition.Text) - 0.01
'End If
'plot
'End Sub

Private Sub txtX_Change()
'txtXX.Text = txtX.Text + txtZoom.Text
plot
End Sub

Private Sub txtY_Change()
plot
End Sub

Private Sub txtZoom_Change()
lblZoom.Caption = "Zoom = " + Format((hsZoom.Value / (hsZoom.Max / 2)) * 100, "##") + "%"
plot
End Sub
