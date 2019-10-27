VERSION 5.00
Begin VB.Form SobreSistema 
   Caption         =   "Sobre o Sistema"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdvoltar 
      Caption         =   "&Voltar"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblinformaçao 
      AutoSize        =   -1  'True
      Caption         =   "Sistema de Cadastro de Funcionários  produzida no Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   120
      Picture         =   "SobreSistema.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "SobreSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdvoltar_Click()
   ProjetoFinal.Show
   Unload Me
End Sub
