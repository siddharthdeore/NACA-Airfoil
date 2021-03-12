VERSION 5.00
Begin VB.Form frmPoints 
   Caption         =   "Cordinates"
   ClientHeight    =   8664
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6132
   Icon            =   "frmPoints.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8664
   ScaleWidth      =   6132
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox txtPoints 
      Height          =   7935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
Dim fname As String
'fname = App.Path + "\" + Format(Now, "DD-MM-YYYY hh.mm.ss") + "Points.txt"
fname = App.Path + "\NACA" + frmCanvas.NACA + ".DAT"
'NACA
Dim fnum As Integer

    ' Open the file for append.
    fnum = FreeFile
    Open fname For Output As fnum
  
    ' Add the command.
    txt = txtPoints.Text
    Print #fnum, txt
    
    ' Close the file.
    Close fnum
    MsgBox "Points saved at " + fname
End Sub
