VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValider 
      Caption         =   "Valider"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSuppr 
      Caption         =   "Supprimer"
      Height          =   975
      Left            =   5400
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ListBox ListPolice 
      Height          =   2010
      ItemData        =   "mainForm.frx":0000
      Left            =   960
      List            =   "mainForm.frx":0002
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtTrans 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton quitter 
      Caption         =   "Quitter"
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblVisu 
      Height          =   855
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label labelAjout 
      Caption         =   "Texte à transcrire :"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub Form_Load()
    ListPolice.AddItem ("MV Boli")
    ListPolice.AddItem ("WolfsRain")
End Sub

Private Sub cmdValider_Click()
    Dim f As New StdFont
    
    If ListPolice = "MV Boli" Then
        f.Name = "MV Boli"
        Set lblVisu.Font = f
        lblVisu = txtTrans
    ElseIf ListPolice = "WolfsRain" Then
        f.Name = "WolfsRain"
        Set lblVisu.Font = f
        lblVisu.FontSize = 20
        lblVisu = txtTrans
    ElseIf ListPolice = "" Then
        lblVisu = txtTrans
    End If
End Sub

Private Sub cmdSuppr_Click()
    ListPolice.RemoveItem (ListPolice.ListIndex)
End Sub

Private Sub quitter_Click()
    Unload Form1
End Sub



