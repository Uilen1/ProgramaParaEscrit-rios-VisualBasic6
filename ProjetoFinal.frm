VERSION 5.00
Begin VB.Form ProjetoFinal 
   Caption         =   "Projeto Final"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdSobre 
      Caption         =   "&Sobre"
      Height          =   375
      Left            =   6240
      TabIndex        =   28
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alt&erar"
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Re&mover"
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Últim&o"
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "Pró&ximo"
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "An&terior"
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrimeiro 
      Caption         =   "&Primeiro"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton cmdsair 
         Caption         =   "Sair"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdpesquisa 
         Caption         =   "Pesquisa"
         Height          =   375
         Left            =   3600
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtSalario 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox TxtId 
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   11
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox TxtFone 
         Height          =   285
         Left            =   4680
         MaxLength       =   9
         TabIndex        =   10
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "ProjetoFinal.frx":0000
         Left            =   2280
         List            =   "ProjetoFinal.frx":0016
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox TxtCidade 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox TxtBairro 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox TxtEnd 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtNomefunc 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtFuncionario 
         Height          =   375
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblSalario 
         AutoSize        =   -1  'True
         Caption         =   "Salário R$:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   18
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Caption         =   "Identidade nº:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   17
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Label lblFone 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   16
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   15
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   14
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   13
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label lblEndereço 
         AutoSize        =   -1  'True
         Caption         =   "Endereço Residencial"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Funcionário"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label LblCodFunc 
         AutoSize        =   -1  'True
         Caption         =   "Código do funcionário"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1980
      End
   End
   Begin VB.Menu mnuregistros 
      Caption         =   "&Registros"
      Begin VB.Menu mnuadicionar 
         Caption         =   "&Adicionar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuremover 
         Caption         =   "Re&mover"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnualterar 
         Caption         =   "Alt&erar"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnutraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnulimpar 
         Caption         =   "&Limpar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuimprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnutraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufechar 
         Caption         =   "&Fechar"
      End
   End
   Begin VB.Menu mnunavegar 
      Caption         =   "&Navegar"
      Begin VB.Menu mnuprimeiro 
         Caption         =   "&Primeiro"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuanterior 
         Caption         =   "An&terior"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuproximo 
         Caption         =   "Pró&ximo"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuultimo 
         Caption         =   "Últim&o"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuajuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnusobre 
         Caption         =   "&Sobre o Sistema"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "ProjetoFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###--- Declarações de Variáveis Globais ---###

Public Dados As Database

Public Tabela_Dados As Recordset

Public Pesquisa As Recordset

'#---OUTRA MANEIRA DE CONECTAR AO BANCO DE DADOS

'Dim DB As ADODB.Connection

'Dim Tabela_Dados As ADODB.Recordset

'Public DBname As String


'###--- Declarações de Funções  ---###
Function Remove_Func()
 
 Tabela_Dados.Delete

   If Tabela_Dados.BOF Then
      Tabela_Dados.MoveFirst
      Else
         Tabela_Dados.MoveLast
      End If
   Tabela_Dados.Requery
   
   Mensagem = MsgBox(" O item foi Removido" & Chr$(13) & _
                     "     Com sucesso! ", vbInformation, "")
                     
End Function
Function Imprimi_Dados()
   
   Printer.Print "Código: "; txtFuncionario.Text
   Printer.Print "Nome: "; txtNomefunc.Text
   Printer.Print "Endereço: "; TxtEnd.Text
   Printer.Print "Bairro: "; TxtBairro.Text
   Printer.Print "Cidade: "; TxtCidade.Text
   Printer.Print "Estado: "; Combo1.Text
   Printer.Print "Telefone: "; TxtFone.Text
   Printer.Print "Registro: "; TxtId.Text
   Printer.Print "Salário: "; TxtSalario.Text
   Printer.EndDoc
   
End Function
Function Atualiza_Campos()

If Tabela_Dados.EOF Or Tabela_Dados.BOF Then
   Limpa_Dados
   Else
   If Tabela_Dados("Cod").Value > 0 Then
         txtFuncionario.Text = Tabela_Dados("Cod")
      Else
         txtFuncionario.Text = ""
   End If
   
   If Tabela_Dados("Nome").Value > 0 Then
         txtNomefunc.Text = Tabela_Dados("Nome")
      Else
         txtNomefunc.Text = ""
   End If
   
   If Tabela_Dados("EndereçoResidencial").Value > 0 Then
         TxtEnd.Text = Tabela_Dados("EndereçoResidencial")
      Else
         TxtEnd.Text = ""
   End If
   
   If Tabela_Dados("Bairro").Value > 0 Then
         TxtBairro.Text = Tabela_Dados("Bairro")
      Else
         TxtBairro.Text = ""
   End If
   
   If Tabela_Dados("Cidade").Value > 0 Then
         TxtCidade.Text = Tabela_Dados("Cidade")
      Else
         TxtCidade.Text = ""
   End If
   
   If Tabela_Dados("Estado").Value > 0 Then
         Combo1.Text = Tabela_Dados("Estado")
      Else
         Combo1.Text = ""
   End If
   
   If Tabela_Dados("Telefone").Value > 0 Then
         TxtFone.Text = Tabela_Dados("Telefone")
      Else
         TxtFone.Text = ""
   End If
   
   If Tabela_Dados("Identidade").Value > 0 Then
         TxtId.Text = Tabela_Dados("Identidade")
      Else
         TxtId.Text = ""
   End If
   
   If Tabela_Dados("Salário").Value > 0 Then
         TxtSalario.Text = Tabela_Dados("Salário")
      Else
         TxtSalario.Text = ""
   End If
End If

End Function

Function Limpa_Dados()

   txtFuncionario.Text = ""
   txtNomefunc.Text = ""
   TxtEnd.Text = ""
   TxtBairro.Text = ""
   TxtCidade.Text = ""
   Combo1.Text = ""
   TxtFone.Text = ""
   TxtId.Text = ""
   TxtSalario.Text = ""
   
End Function
Function Adicionar_Func()
   
   Tabela_Dados.AddNew
   
   If txtFuncionario.Text = "" Then
      txtFuncionario.Text = " "
   End If
   
   If txtNomefunc.Text = "" Then
      txtNomefunc.Text = " "
   End If
   
   If TxtEnd.Text = "" Then
      TxtEnd.Text = " "
   End If
   
   If TxtBairro.Text = "" Then
      TxtBairro.Text = " "
   End If
   
   If TxtCidade.Text = "" Then
      TxtCidade.Text = " "
   End If
   
   If Combo1.Text = "" Then
      Combo1.Text = " "
   End If
   
   If TxtFone.Text = "" Then
      TxtFone.Text = " "
   End If
   
   If TxtId.Text = "" Then
      TxtId.Text = " "
   End If
   
   If TxtSalario.Text = "" Then
      TxtSalario.Text = " "
   End If
   
   Tabela_Dados("Cod") = txtFuncionario.Text
   Tabela_Dados("Nome") = txtNomefunc.Text
   Tabela_Dados("EndereçoResidencial") = TxtEnd.Text
   Tabela_Dados("Bairro") = TxtBairro.Text
   Tabela_Dados("Cidade") = TxtCidade.Text
   Tabela_Dados("Estado") = Combo1.Text
   Tabela_Dados("Telefone") = TxtFone.Text
   Tabela_Dados("Identidade") = TxtId.Text
   Tabela_Dados("Salário") = TxtSalario.Text
   Tabela_Dados.Update
   Tabela_Dados.Requery
   Mensagem = MsgBox(" O Funcionário foi adicionado" & Chr$(13) & _
                     "          Com sucesso!", vbInformation, "")

End Function
Function Alterar_Func()
   
   Tabela_Dados.Edit
   Tabela_Dados("Cod") = txtFuncionario.Text
   Tabela_Dados("Nome") = txtNomefunc.Text
   Tabela_Dados("EndereçoResidencial") = TxtEnd.Text
   Tabela_Dados("Bairro") = TxtBairro.Text
   Tabela_Dados("Cidade") = TxtCidade.Text
   Tabela_Dados("Estado") = Combo1.Text
   Tabela_Dados("Telefone") = TxtFone.Text
   Tabela_Dados("Identidade") = TxtId.Text
   Tabela_Dados("Salário") = TxtSalario.Text
   Mensagem = MsgBox(" O Item foi Alterado" & Chr$(13) & _
                     "    Com sucesso!", vbInformation, "")
   Tabela_Dados.Update
   Tabela_Dados.Requery
   
   End Function
'###--- Programando os eventos dos Objetos ---###

Private Sub cmdAdd_Click()

If Tabela_Dados.RecordCount = 0 Then
   cmdRemover.Enabled = True
   cmdAlterar.Enabled = True
End If
   Adicionar_Func
   Limpa_Dados
   
End Sub

Private Sub cmdAlterar_Click()

If Tabela_Dados.RecordCount = 0 Then
   Mensagem = MsgBox("Não há Registros!", vbOKOnly, "")
   cmdAlterar.Enabled = False
   Exit Sub
   Else
   Alterar_Func
End If

End Sub

Private Sub cmdAnterior_Click()

If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MovePrevious
   
   If Tabela_Dados.BOF Then
      Tabela_Dados.MoveFirst
   End If
End If
   Atualiza_Campos
   
End Sub

Private Sub cmdFechar_Click()
   End
End Sub

Private Sub cmdImprimir_Click()
   Imprimi_Dados
End Sub

Private Sub cmdLimpar_Click()
   Limpa_Dados
End Sub

Private Sub cmdpesquisa_Click()
      
   Dim Dado, Consulta, Mensagem, Busca As String

   Dado = InputBox(" Digite o Nome: ", "Pesquisa de Funcionários")
   
   Consulta = "select * from Dados where nome like '" + Dado + "*' order by nome;"
   
   Set Tabela_Dados = Dados.OpenRecordset(Consulta, dbOpenDynaset)
   
   If Tabela_Dados.BOF Or Tabela_Dados.EOF Then
      MsgBox ("Nome não encontrado!")
   End If
   Tabela_Dados.Requery
   Atualiza_Campos
   cmdsair.Enabled = True
End Sub

Private Sub cmdPrimeiro_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveFirst
End If
   Atualiza_Campos
   
End Sub

Private Sub cmdProximo_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveNext
   
   If Tabela_Dados.EOF Then
      Tabela_Dados.MoveLast
   End If
End If
   Atualiza_Campos
End Sub

Private Sub cmdRemover_Click()
If Tabela_Dados.RecordCount < 1 Then
   cmdRemover.Enabled = False
   Exit Sub
   Else
      Remove_Func
      Limpa_Dados
End If
Tabela_Dados.Requery
End Sub

Private Sub cmdsair_Click()
   Set Dados = OpenDatabase(App.Path & "\dados.mdb", _
   False)
   Set Tabela_Dados = Dados.OpenRecordset("Dados", _
   dbOpenDynaset)
   MsgBox ("Saindo da Pesquisa!")
   Limpa_Dados
   cmdsair.Enabled = False
End Sub

Private Sub cmdSobre_Click()
   SobreSistema.Show
   
End Sub

Private Sub cmdUltimo_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveLast
End If
   Atualiza_Campos

End Sub
'###--- Definindo o Carregamento do Form Principal ---###
Private Sub Form_Load()

   Set Dados = OpenDatabase(App.Path & "\dados.mdb", _
   False)
   Set Tabela_Dados = Dados.OpenRecordset("Dados", _
   dbOpenDynaset)
   
   
   Tabela_Dados.AddNew
   
   Atualiza_Campos
'##----OUTRA MANEIRA DE CONECTAR O BANCO DE DADOS
   'If Right(App.Path, 1) = "\" Then
  'DBname = App.Path & "Dados.mdb"
    'Else
  'DBname = App.Path & "\" & "Dados.mdb"
   'End If
   
   'Set DB = New ADODB.Connection
   'Set Tabela_Dados = New ADODB.Recordset
   
   'DB.Mode = adModeReadWrite
   
   'DB.Open " Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBname & ";"
   
   'Tabela_Dados.Open "select * from dados order by Nome ASC", DB, adOpenStatic, adLockOptimistic
End Sub

'###--- Programando os eventos dos Menus ---###
Private Sub mnuadicionar_Click()
   If Tabela_Dados.RecordCount = 0 Then
   cmdRemover.Enabled = True
   cmdAlterar.Enabled = True
End If
   Adicionar_Func
End Sub

Private Sub mnualterar_Click()
   If Tabela_Dados.RecordCount = 0 Then
   Mensagem = MsgBox("Não há Registros!", vbOKOnly, "")
   cmdAlterar.Enabled = False
   Exit Sub
   Else
   Alterar_Func
End If
End Sub

Private Sub mnuanterior_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MovePrevious
   
   If Tabela_Dados.BOF Then
      Tabela_Dados.MoveFirst
   End If
End If
   Atualiza_Campos
   
End Sub

Private Sub mnufechar_Click()
   End
End Sub

Private Sub mnuimprimir_Click()
   Imprimi_Dados
End Sub

Private Sub mnulimpar_Click()
   Limpa_Dados
End Sub

Private Sub mnuprimeiro_Click()
   If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveFirst
End If
   Atualiza_Campos
End Sub

Private Sub mnuproximo_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveNext
   
   If Tabela_Dados.EOF Then
      Tabela_Dados.MoveLast
   End If
End If
   Atualiza_Campos
End Sub

Private Sub mnuremover_Click()
   If Tabela_Dados.RecordCount < 1 Then
   cmdRemover.Enabled = False
   Exit Sub
   Else
      Remove_Func
      Limpa_Dados
End If
Tabela_Dados.Requery
End Sub

Private Sub mnusobre_Click()
   SobreSistema.Show
End Sub

Private Sub mnuultimo_Click()
If Tabela_Dados.RecordCount <= 0 Then
   Exit Sub
   Else
   Tabela_Dados.MoveLast
End If
   Atualiza_Campos
End Sub

