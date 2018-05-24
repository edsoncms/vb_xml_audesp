VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MS3TI - AUDESP XML"
   ClientHeight    =   7290
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   10755
   Icon            =   "audesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox carrega 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   3480
      Picture         =   "audesp.frx":424A
      ScaleHeight     =   1695
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Selecionar Arquivo Histórico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Gerar XML"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   3
      Top             =   6240
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Selecionar Arquivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo de Arquivo"
      Height          =   1215
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   4815
      Begin VB.OptionButton optlotacao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lotação"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   3855
      End
      Begin VB.OptionButton optpublico 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agente Público"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "SISTEMA GERAÇÃO DE XML PARA A AUDESP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   3645
      Left            =   0
      Picture         =   "audesp.frx":F288
      Top             =   120
      Width           =   4995
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   8760
      Picture         =   "audesp.frx":12FF8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label nomearquivo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   10215
   End
   Begin VB.Label nomearquivo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.Menu sobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 Dim ff As Integer
ff = FreeFile 'Sets to next available file number
With CommonDialog1
    .FileName = ""
    .Filter = "All files (*.*) |*.*|" 'Sets the filter
    .ShowOpen
End With
nomearquivo.Caption = CommonDialog1.FileName
nomearquivo.Visible = True

End Sub

Private Sub Command2_Click()


If nomearquivo.Caption = "" Then

MsgBox "Selecione um arquivo primeiro", vbCritical + vbOKOnly

Exit Sub

End If


If optpublico.Value = False And optlotacao.Value = False Then

MsgBox "Selecione um tipo de arquivo ", vbCritical + vbOKOnly

Exit Sub

End If


If optlotacao.Value = True Then


If nomearquivo.Caption = "" Or nomearquivo2.Caption = "" Then

MsgBox "Selecione os dois arquivos a serem processados.", vbCritical + vbOKOnly
Exit Sub
End If
End If


If MsgBox("Deseja realmente iniciar o processamento?", vbQuestion + vbYesNoCancel) = vbYes Then


    ' MsgBox "Passou", vbInformation + vbOKOnly

    carrega.Visible = True
    
    
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim data_hoje As String
    Dim valor_data As Variant
    
    
    data_hoje = Format(Date, "yyyy-mm-dd")

    

    
    'MsgBox data_hoje



    Dim F               As Long
    Dim Linha           As String
    Dim i               As Long
    Dim Tmp             As String
    Dim texto           As String
    
    
    If optpublico.Value = True Then
    
    
    Dim nome_agente      As String
    Dim cpf              As String
    Dim pis              As String
    Dim data_nascimento  As String
    Dim sexo             As String
    Dim nacionalidade    As String
    Dim escolaridade     As String
    Dim especialidade    As String


    ElseIf optlotacao.Value = True Then
    
    
    Dim cpf2                 As String
    Dim data_lotacao2        As String
    Dim nome_agente2         As String
    Dim exercicio_atividade  As String
    Dim codigo_nome_cargo    As String
    Dim codigo_nome_funcao   As String
    Dim cargo_funcao_remun   As String
    Dim unidade_lotacao      As String
    Dim funcao_governo       As String
    Dim forma_provimento     As String
    Dim data_exercicio       As String
    

    Dim cpf_historico            As String
    Dim cargo_funcao_his         As String
    Dim data_lotacao_his         As String
    Dim nome_agente_his          As String
    Dim data_his_situacoes       As String
    Dim st_his_situacoes         As String
    
    
    
    
    End If




    F = FreeFile
            Open nomearquivo.Caption For Input As #F
            
            

    
    If optpublico.Value = True Then
    
    texto = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" + vbCrLf
    texto = texto + "<ag:AgentesPublicos xsi:schemaLocation=""http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-agentepublico ../quadrofuncional/AUDESP_AGENTEPUBLICO_2016_A.XSD"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:ag=""http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-agentepublico"" xmlns:ap=""http://www.tce.sp.gov.br/audesp/xml/atospessoal"" xmlns:gen=""http://www.tce.sp.gov.br/audesp/xml/generico"">" + vbCrLf
    
    
    
    
                texto = texto + "<ag:Descritor>   " + vbCrLf
                texto = texto + "<gen:AnoExercicio>2016</gen:AnoExercicio> " + vbCrLf
                texto = texto + "<gen:TipoDocumento>Agente Público</gen:TipoDocumento> " + vbCrLf
                texto = texto + "<gen:Entidade>10063</gen:Entidade> " + vbCrLf
                texto = texto + "<gen:Municipio>7107</gen:Municipio> " + vbCrLf
                texto = texto + "<gen:DataCriacaoXML>" + data_hoje + "</gen:DataCriacaoXML> " + vbCrLf
                texto = texto + "</ag:Descritor> " + vbCrLf
                texto = texto + "<ag:ListaAgentePublico>" + vbCrLf
    End If
    
    If optlotacao.Value = True Then
    
    
                texto = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" + vbCrLf
                texto = texto + "<LotacaoAgentePublico xsi:schemaLocation=""http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-lotacaoagentepublico ../quadrofuncional/AUDESP_LOTACAOAGENTEPUBLICO_2016_A.XSD"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-lotacaoagentepublico"" xmlns:qpla=""http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-lotacaoagentepublico"" xmlns:ap=""http://www.tce.sp.gov.br/audesp/xml/atospessoal"" xmlns:gen=""http://www.tce.sp.gov.br/audesp/xml/generico"">" + vbCrLf
    
                texto = texto + "<Descritor>" + vbCrLf
                texto = texto + "<gen:AnoExercicio>2016</gen:AnoExercicio>" + vbCrLf
                texto = texto + "<gen:TipoDocumento>Lotação Agente Público</gen:TipoDocumento>" + vbCrLf
                texto = texto + "<gen:Entidade>10063</gen:Entidade>" + vbCrLf
                texto = texto + "<gen:Municipio>7107</gen:Municipio>" + vbCrLf
                texto = texto + "<gen:DataCriacaoXML>" + data_hoje + "</gen:DataCriacaoXML>" + vbCrLf
                texto = texto + "</Descritor>" + vbCrLf
                
                texto = texto + "<ListaLotacaoAgentePublico>" + vbCrLf
    
    End If
    


    Do While Not EOF(F)
    Line Input #F, Linha 'lê uma linha do arquivo texto


            testpos = InStr(Linha, ";")
            ' ///////////////VERIFICA SE A LINHA TEM PONTO E VÍRGULA , SE SIM SIGNIFICA QUE TEM DADOS PARA SEREM LIDOS
            If testpos <> 0 Then
                Var = Split(Linha, ";")
                
                ' CASO TENHA ESCOLHIDO A OPÇÃO AGENTE PÚBLICO
                If optpublico.Value = True Then


                nome_agente = Var(0)
                cpf = Format(Var(1), "00000000000")
                pis = Format(Var(2), "00000000000")
                data_nascimento = Format(Var(3), "yyyy-mm-dd")
                sexo = Var(4)
                nacionalidade = Var(5)
                escolaridade = Var(6)
                especialidade = Var(7)

                

                texto = texto + "<ag:AgentePublico>" + vbCrLf
                texto = texto + "<ag:nome>" + nome_agente + "</ag:nome>" + vbCrLf
                texto = texto + "<ag:cpf Tipo=""02"">" + vbCrLf
                texto = texto + "<gen:Numero>" + cpf + "</gen:Numero>" + vbCrLf
                texto = texto + "</ag:cpf>" + vbCrLf
                texto = texto + "<ag:pis_pasep>" + pis + "</ag:pis_pasep>" + vbCrLf
                texto = texto + "<ag:dataNascimento>" + data_nascimento + "</ag:dataNascimento>" + vbCrLf
                texto = texto + "<ag:sexo>" + sexo + "</ag:sexo>" + vbCrLf
                texto = texto + "<ag:nacionalidade>" + nacionalidade + "</ag:nacionalidade>" + vbCrLf
                texto = texto + "<ag:escolaridade>" + escolaridade + "</ag:escolaridade>" + vbCrLf
                texto = texto + "<ag:especialidade>" + especialidade + "</ag:especialidade>" + vbCrLf
                texto = texto + "</ag:AgentePublico>" + vbCrLf
                End If
                
                
                
                ' CASO TENHA ESCOLHIDO A OPÇÃO AGENTE PÚBLICO
                If optlotacao.Value = True Then


                cpf = Format(Var(0), "00000000000")
                data_lotacao = Format(Var(1), "yyyy-mm-dd")
                nome_agente = Var(2)
                exercicio_atividade = Var(3)
                codigo_nome_cargo = Var(4)
                codigo_nome_funcao = Var(5)
                cargo_funcao = Var(6)
                unidade_lotacao = Var(7)
                funcao_governo = Var(8)
                forma_provimento = Var(9)
               

                

                texto = texto + "<LotacaoAgentePublico>" + vbCrLf
                texto = texto + "<cpf Tipo=""02"">" + vbCrLf
                texto = texto + "<gen:Numero>" + cpf + "</gen:Numero>" + vbCrLf
                texto = texto + "</cpf>" + vbCrLf
                texto = texto + "<dataLotacao>" + data_lotacao + "</dataLotacao>" + vbCrLf
                texto = texto + "<exercicioAtividade>" + exercicio_atividade + "</exercicioAtividade>" + vbCrLf
                texto = texto + "<codigoMunicipioCargo>6101</codigoMunicipioCargo>" + vbCrLf
                texto = texto + "<codigoEntidadeCargo>02</codigoEntidadeCargo>" + vbCrLf
                texto = texto + "<codigoCargo>" + codigo_nome_cargo + "</codigoCargo>" + vbCrLf
                texto = texto + "<cargoRemunerado/>" + vbCrLf
                texto = texto + "<unidadeLotacao>" + unidade_lotacao + "</unidadeLotacao>" + vbCrLf
                texto = texto + "<funcaoGoverno>" + funcao_governo + "</funcaoGoverno>" + vbCrLf
                texto = texto + "<formaProvimento>" + forma_provimento + "</formaProvimento>" + vbCrLf
                texto = texto + "<dataExercicio>" + data_lotacao + "</dataExercicio>" + vbCrLf
                texto = texto + "</LotacaoAgentePublico>" + vbCrLf
                
             End If
             '   MsgBox texto, vbInformation + vbOKOnly
            End If
            
                              
                                              
    Loop
            If optpublico.Value = True Then
                texto = texto + "</ag:ListaAgentePublico>" + vbCrLf
                texto = texto + "</ag:AgentesPublicos>" + vbCrLf
            End If
            
            

            If optlotacao.Value = True Then
                texto = texto + "</ListaLotacaoAgentePublico>" + vbCrLf
            End If
            
            If optlotacao.Value = True Then
            
           G = FreeFile
            Open nomearquivo2.Caption For Input As #G
            
            
           texto = texto + "<ListaHistoricoLotacaoAgentePublico>" + vbCrLf
            
            
            Do While Not EOF(G)
            Line Input #G, Linha 'lê uma linha do arquivo texto


            testpos2 = InStr(Linha, ";")
            ' ///////////////VERIFICA SE A LINHA TEM PONTO E VÍRGULA , SE SIM SIGNIFICA QUE TEM DADOS PARA SEREM LIDOS
            If testpos2 <> 0 Then
                var2 = Split(Linha, ";")
            
            
            
            
            cpf_historico = var2(0)
            cargo_funcao_his = var2(1)
            data_lotacao_his = Format(var2(2), "yyyy-mm-dd")
            nome_agente_his = var2(3)
            data_his_situacoes = Format(var2(4), "yyyy-mm-dd")
            st_his_situacoes = var2(5)
            
            
            

            
             texto = texto + "<HistoricoLotacaoAgentePublico>" + vbCrLf
             texto = texto + "<cpf Tipo=""02"">" + vbCrLf
             texto = texto + "<gen:Numero>" + cpf_historico + "</gen:Numero>" + vbCrLf
             texto = texto + "</cpf>" + vbCrLf
             texto = texto + "<dataExercicio>" + data_lotacao_his + "</dataExercicio>" + vbCrLf
             texto = texto + "<dataLotacao>" + data_lotacao_his + "</dataLotacao>" + vbCrLf
             texto = texto + "<codigoMunicipioFuncao>6101</codigoMunicipioFuncao>" + vbCrLf
             texto = texto + "<codigoEntidadeFuncao>02</codigoEntidadeFuncao>" + vbCrLf
             texto = texto + "<codigoFuncao>" + codigo_nome_funcao + "</codigoFuncao>" + vbCrLf
             texto = texto + "<dataHistoricoLotacao>" + data_lotacao_his + "</dataHistoricoLotacao>" + vbCrLf
             texto = texto + "<situacao>4</situacao>" + vbCrLf
             texto = texto + "<codigoMunicipioLotacao>0000</codigoMunicipioLotacao>" + vbCrLf
             texto = texto + "<codigoEntidadeLotacao>01</codigoEntidadeLotacao>" + vbCrLf
             texto = texto + "</HistoricoLotacaoAgentePublico>" + vbCrLf
            
            
            
            End If
            
            
             Loop
             
             Close #G
           End If
            
            

            If optlotacao.Value = True Then
                texto = texto + "</ListaHistoricoLotacaoAgentePublico>" + vbCrLf
                texto = texto + "</LotacaoAgentePublico>" + vbCrLf
            End If
            
            
            



            Close #F
            
            If optpublico.Value = True Then

             Open App.Path & "/AgentePublico.xml" For Binary As #1
             Put #1, , texto
             Close #1
             
            End If
            
            If optlotacao.Value = True Then
            
             Open App.Path & "/Lotacao.xml" For Binary As #1
             Put #1, , texto
             Close #1
            
            
            End If
           


           



    carrega.Visible = False
    If optpublico.Value = True Then
    
    MsgBox "Arquivo gerado com sucesso no local onde este programa está salvo com o nome de AgentePublico.xml!", vbInformation + vbOKOnly

    End If

    If optlotacao.Value = True Then
    
    MsgBox "Arquivo gerado com sucesso no local onde este programa está salvo com o nome de Lotacao.xml!", vbInformation + vbOKOnly

    End If










End If




End Sub



Private Sub Command3_Click()
 Dim ff As Integer
ff = FreeFile 'Sets to next available file number
With CommonDialog1
    .FileName = ""
    .Filter = "All files (*.*) |*.*|" 'Sets the filter
    .ShowOpen
End With
nomearquivo2.Caption = CommonDialog1.FileName
nomearquivo2.Visible = True
End Sub

Private Sub optlotacao_Click()
Command3.Visible = True
End Sub

Private Sub optpublico_Click()
Command3.Visible = False
End Sub

Private Sub sobre_Click()
    frmAbout.Show

End Sub
