unit UClientes;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Datasnap.DBClient, System.StrUtils,
  Vcl.DBCtrls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Vcl.Mask,
  sLabel, Vcl.Grids, Vcl.DBGrids, System.Actions, Vcl.ActnList,
  System.ImageList, Vcl.ImgList, Datasnap.Provider, Data.FMTBcd, Data.SqlExpr,
  acPNG, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, MemDS, DBAccess, VirtualQuery,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.FileCtrl, Vcl.ComCtrls,
  System.json, IdHTTP, IdSSLOpenSSL,

  REST.Client, IPPeerClient;

  procedure Envia_Email( cPortaSMTP_Padrao, cHostSMTP_Padrao, cUserNameSMTP_Padrao, cSenhaUserSMTP_Padrao,
                              cEmailOrigem, cName, cEmailDestino, cAssuntoEmail, cCorpoEmail, cCaminhoArqAnexo, cAlertaEnvioEmail : WideString); stdcall;
  external 'LibMatrix.dll' name 'Envia_Email';

type
  TFrmClientes = class(TForm)
    CDS_Clientes: TClientDataSet;
    DS_Clientes: TDataSource;
    ActList_Acoes: TActionList;
    ActIncluir: TAction;
    ImageList1: TImageList;
    ActAlterar: TAction;
    ActSalvar: TAction;
    ActExcluir: TAction;
    DSP_Cliente: TDataSetProvider;
    CDS_ClientesID: TIntegerField;
    CDS_ClientesNOME: TStringField;
    CDS_ClientesRG: TStringField;
    CDS_ClientesCPF: TStringField;
    CDS_ClientesFONE: TStringField;
    CDS_ClientesEMAIL: TStringField;
    CDS_ClientesENDERECO: TStringField;
    CDS_ClientesCEP: TStringField;
    CDS_ClientesLOGRADOURO: TStringField;
    CDS_ClientesNUMERO: TIntegerField;
    CDS_ClientesCOMPLEMENTO: TStringField;
    CDS_ClientesBAIRRO: TStringField;
    CDS_ClientesCIDADE: TStringField;
    CDS_ClientesESTADO: TStringField;
    CDS_ClientesPAIS: TStringField;
    Image16: TImage;
    PageControl1: TPageControl;
    TabSheet_Dados: TTabSheet;
    GroupBox_Dados: TGroupBox;
    Image_AltP_PesqProd: TImage;
    Image4: TImage;
    Image6: TImage;
    Image11: TImage;
    DBText1: TDBText;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    sLabel1: TsLabel;
    Image1: TImage;
    Image5: TImage;
    Label1: TLabel;
    Image7: TImage;
    Label5: TLabel;
    Label7: TLabel;
    Image8: TImage;
    Label8: TLabel;
    Image9: TImage;
    Label9: TLabel;
    Image10: TImage;
    Label10: TLabel;
    Label13: TLabel;
    Image12: TImage;
    Label14: TLabel;
    Image13: TImage;
    Label15: TLabel;
    Image14: TImage;
    Label16: TLabel;
    Image15: TImage;
    DBEdit_Nome: TDBEdit;
    DBEdit_Identidade: TDBEdit;
    DBEdit_Email: TDBEdit;
    DBEdit_Telefone: TDBEdit;
    DBEdit_CPF: TDBEdit;
    DBEdit_CEP: TDBEdit;
    DBEdit_Endereco: TDBEdit;
    DBEdit_Logradouro: TDBEdit;
    DBEdit_Numero: TDBEdit;
    DBEdit_Complemento: TDBEdit;
    DBEdit_Bairro: TDBEdit;
    DBEdit_Cidade: TDBEdit;
    DBComboBox_UF: TDBComboBox;
    DBComboBox_Pais: TDBComboBox;
    Panel7: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    Image2: TImage;
    Image3: TImage;
    ComboBox_Clientes: TComboBox;
    BitBtn_Localizar: TBitBtn;
    Panel_Registros: TPanel;
    ComboBox_Pesquisa: TComboBox;
    DBNavigator_Clientes: TDBNavigator;
    BitBtn_Inclui: TBitBtn;
    BitBtn_Altera: TBitBtn;
    BitBtn_Salvar: TBitBtn;
    BitBtn_Excluir: TBitBtn;
    BitBtn_Sair: TBitBtn;
    DBGrid_Clientes: TDBGrid;
    TabSheet_EnvioEmail: TTabSheet;
    FileListBox1: TFileListBox;
    DirectoryListBox1: TDirectoryListBox;
    DriveComboBox1: TDriveComboBox;
    Edit_CaminhoSelecionado: TEdit;
    Label17: TLabel;
    Label18: TLabel;
    Spb_EnviaEmail: TSpeedButton;
    GroupBox1: TGroupBox;
    Label19: TLabel;
    Edit_PortaSMTP: TEdit;
    Label20: TLabel;
    Edit_ServidorSMTP: TEdit;
    Label21: TLabel;
    Edit_UsuarioSMTP: TEdit;
    Label22: TLabel;
    Edit_SenhaSMTP: TEdit;
    BitBtn_CarregaXML: TBitBtn;
    BitBtn_BuscaCEP: TBitBtn;
    BitBtn_SalvaXML: TBitBtn;
    Image17: TImage;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ActIncluirExecute(Sender: TObject);

    procedure Checa_Botoes;
    procedure Painel_Registros;
    function Valida_Campos : Boolean;
    procedure Habilita_Controles( bHabilita : Boolean );

    procedure BitBtn_SairClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DBEdit_NomeKeyPress(Sender: TObject; var Key: Char);
    procedure DBGrid_ClientesTitleClick(Column: TColumn);
    procedure ActAlterarExecute(Sender: TObject);
    procedure ActSalvarExecute(Sender: TObject);
    procedure ActExcluirExecute(Sender: TObject);

    procedure Proced_EntraemDigitacao( Sender : TObject );
    procedure Proced_SaidaDigitacao( Sender : TObject );

    procedure EntraemDigitacao( Sender : TObject );
    procedure SaidaDigitacao( Sender : TObject );
    procedure Spb_EnviaEmailClick(Sender: TObject);
    procedure DirectoryListBox1DblClick(Sender: TObject);
    procedure DriveComboBox1Change(Sender: TObject);
    procedure DirectoryListBox1Change(Sender: TObject);
    procedure TabSheet_EnvioEmailEnter(Sender: TObject);
    procedure BitBtn_LocalizarClick(Sender: TObject);
    procedure ComboBox_PesquisaKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBox_ClientesChange(Sender: TObject);
    procedure ComboBox_ClientesKeyPress(Sender: TObject; var Key: Char);
    procedure BitBtn_CarregaXMLClick(Sender: TObject);
    procedure BitBtn_BuscaCEPClick(Sender: TObject);
    procedure BitBtn_SalvaXMLClick(Sender: TObject);

  private
    { Private declarations }
    function ChecaCPF(CPF : string; bAlerta : Boolean) : Boolean;
    function ValidaEmail(const s:string): Boolean;
    function Replicate( pString:String; xWidth:Integer ) :String;
    function Captura_SoNumeroString( cTexto: string ) : string;

    function BuscarCEP_ViaCEP(UmCEP: string): TStringList;

  public
    { Public declarations }

  end;

var
  FrmClientes: TFrmClientes;

  aVet_Sequencia : array[0..9] of string = ('00000000000',
                                            '11111111111',
                                            '22222222222',
                                            '33333333333',
                                            '44444444444',
                                            '55555555555',
                                            '66666666666',
                                            '77777777777',
                                            '88888888888',
                                            '99999999999' );

  //-------------
  // Padrão de cor de fundo e de letra para função
  // "EntraEmDigitacao" e "SaiDigitacao"
  //----------------------------------------

  Entra_Digitacao_Cor_Fundo  : TColor = clAqua;
  Entra_Digitacao_Cor_Letra  : TColor = clBlue;

  Sai_Digitacao_Cor_Fundo    : TColor = $00D5FFAD; // verde levemente claro
  Sai_Digitacao_Cor_Letra    : TColor = clBlack;

  // Usado esta cor de fundo no Sistema de Ordem de Serviço (Tela de OS)...
  Sai_Digitacao_Cor_Fundo2   : TColor = clInfobk;
  Sai_Digitacao_Cor_Letra2   : TColor = clMaroon;

  cCorFundo_ANT, cCorLetra_ANT : TColor;

  Autopreencher                : Boolean;
implementation

{$R *.dfm}

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ActAlterarExecute(Sender: TObject);
begin

  If (ComboBox_Clientes.focused) then
    begin
      MessageBeep(0);
      MessageDlg('Atenção...Impossível alterar Cliente !'+#13+#13+'Conclua a pesquisa primeiro ou cancele !', mtError, [mbOk], 0 );
      Exit;
    end;

  MessageBeep(16);
  If MessageDlg('ALTERAR, este Cliente ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
    begin

      Try
        BitBtn_Inclui.enabled        := False;
        BitBtn_Altera.enabled        := False;
        BitBtn_Salvar.enabled        := True;
        BitBtn_Excluir.enabled       := False;
        BitBtn_Localizar.enabled     := False;
        BitBtn_Sair.enabled          := False;
        DBNavigator_Clientes.enabled := False;
        BitBtn_CarregaXML.enabled    := False;
        BitBtn_SalvaXML.enabled      := False;

        Habilita_Controles(True);
        ActSalvar.Enabled            := True;

        CDS_Clientes.Edit;

        If (DBEdit_Nome.CanFocus) then
          DBEdit_Nome.SetFocus;

      Except
        on ERRO : Exception do
          begin
            MessageBeep(0);
            MessageDlg('Atenção...'+#13+#13+'Não foi possivel inciar a alteração deste Cliente !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);

            Application.Terminate;
          end;

      End;

    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ActExcluirExecute(Sender: TObject);
begin

  If (Combobox_Clientes.focused) then
    begin
      MessageBeep(0);
      MessageDlg('Atenção...Impossível excluir Cliente !'+#13+#13+'Conclua a pesquisa primeiro ou cancele !', mtError, [mbOk], 0 );
      Exit;
    end;

  MessageBeep(16);
  If MessageDlg('EXCLUIR, este Cliente ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
    begin
      Try
        CDS_Clientes.Delete;

        Checa_Botoes();
        Painel_Registros();

        //--

        MessageBeep(0);
        MessageDlg('Cliente excluído com sucesso !', mtInformation, [mbOK], 0);

        if BitBtn_Inclui.CanFocus then
          BitBtn_Inclui.SetFocus;

      Except
        on ERRO : Exception do
          begin
            MessageBeep(0);
            MessageDlg('Atenção...'+#13+#13+'Não foi possivel a exclusão deste cliente !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);

            Application.Terminate;
          end;

    End;

    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ActIncluirExecute(Sender: TObject);
Var
  nCodigo : integer;
begin

  If (Combobox_Clientes.focused) then
    begin
      MessageBeep(0);
      MessageDlg('Atenção...Impossível incluir Cliente !'+#13+#13+'Conclua a pesquisa primeiro ou cancele !', mtError, [mbOk], 0 );
      Exit;
    end;

  MessageBeep(16);
  If MessageDlg('INCLUIR, novo Cliente ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
    begin

      Try
        BitBtn_Inclui.enabled        := False;
        BitBtn_Altera.enabled        := False;
        BitBtn_Salvar.enabled        := True;
        BitBtn_Excluir.enabled       := False;
        BitBtn_Localizar.enabled     := False;
        BitBtn_Sair.enabled          := False;
        DBNavigator_Clientes.enabled := False;
        BitBtn_CarregaXML.enabled    := False;
        BitBtn_SalvaXML.enabled      := False;

        Habilita_Controles(True);

        CDS_Clientes.Last;
        nCodigo := CDS_Clientes.FieldByName('ID').AsInteger;
        Inc(nCodigo);

        CDS_Clientes.Append;
        CDS_Clientes.FieldByName('ID').AsInteger  := nCodigo;
        CDS_Clientes.FieldByName('PAIS').AsString := 'Brasil';

        If (DBEdit_Nome.CanFocus) then
          DBEdit_Nome.SetFocus;

        ActSalvar.Enabled          := True;

      Except
        on ERRO : Exception do
          begin
            MessageBeep(0);
            MessageDlg('Atenção...'+#13+#13+'Não foi possivel inciar a inclusão deste Cliente !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);

            Application.Terminate;
          end;

      End;
    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ActSalvarExecute(Sender: TObject);
begin

  if ( not(Valida_Campos()) ) then  Exit;

  //--

  Messagebeep(16);
  If MessageDlg('SALVAR, este Cliente ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
    begin
      Try
        CDS_Clientes.Post;
        //CDS_Clientes.ApplyUpdates(0);

        ActSalvar.Enabled   := False;

        Painel_Registros();
      Except
        on ERRO : Exception do
          begin
            MessageBeep(0);
            MessageDlg('Atenção...'+#13+#13+'Não foi possível salvar este Cliente !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);

            Application.Terminate;
          end;
      End;

      //--

      Habilita_Controles(False);
      Checa_Botoes;

      If BitBtn_Inclui.CanFocus then
        BitBtn_Inclui.SetFocus;

    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.FormCreate(Sender: TObject);
begin

  If PageControl1.ActivePageIndex <> 0 then
    PageControl1.ActivePageIndex := 0;

  Edit_PortaSMTP.Text    := '';
  Edit_ServidorSMTP.Text := '';
  Edit_UsuarioSMTP.Text  := '';
  Edit_SenhaSMTP.Text    := '';

  //--

  CDS_Clientes.CreateDataSet;

  // Abre a tabela depois de criada.
  CDS_Clientes.Open;

  Checa_Botoes;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  If ( ( Shift = [ssAlt] ) And ( Key = VK_F4 ) ) Or ( ( Shift = [ssCtrl] ) And ( Key = VK_F4 ) ) then
    begin
      MessageBeep(0);
      MessageDlg('Procedimento padrão do Windows cancelado !'+#13+#13+'"Feche" a janela usando <ESC> ou o botão de saída !', mtError, [mbOk], 0 );
      Key := 0;
      Exit;
    end;

  //----------------
  // Tratamento de teclas de movimentação...
  //--------------------------------------------

  If (Key = VK_NEXT) And Not( (CDS_Clientes.state = dsEdit) Or (CDS_Clientes.state = dsInsert) )
  and Not(Combobox_Clientes.focused) and Not(Combobox_Pesquisa.focused) then
    begin
      If CDS_Clientes.Bof = True then
        begin
          MessageBeep(16);
          ShowMessage('Você está no primeiro registro do arquivo....');
        end
      Else
        begin
          CDS_Clientes.Prior;

        end;
      Checa_Botoes;
      Painel_Registros();
    end;

  If (Key = VK_PRIOR) And Not( (CDS_Clientes.state = dsEdit) Or (CDS_Clientes.state = dsInsert) )
  and Not(Combobox_Clientes.focused) and Not(Combobox_Pesquisa.focused) then
    begin
      If CDS_Clientes.Eof = True then
        begin
          MessageBeep(16);
          ShowMessage('Você está no último registro do arquivo....');
        end
      Else
        begin
          CDS_Clientes.Next;

        end;

      Checa_Botoes;
      Painel_Registros();
    end;

  If (Key = VK_HOME) And Not( (CDS_Clientes.state = dsEdit) Or (CDS_Clientes.state = dsInsert) )
  and Not(Combobox_Clientes.focused) and Not(Combobox_Pesquisa.focused) then
    begin
      If CDS_Clientes.Bof = True then
        begin
          MessageBeep(16);
          ShowMessage('Você já está no primeiro registro do arquivo....');
        end
      Else
        begin
          CDS_Clientes.First;

        end;

      Checa_Botoes;
      Painel_Registros();
    end;


  If (Key = VK_END) And Not( (CDS_Clientes.state = dsEdit) Or (CDS_Clientes.state = dsInsert) )
  and Not(Combobox_Clientes.focused) and Not(Combobox_Pesquisa.focused) then
    begin
      If CDS_Clientes.Eof = True then
        begin
          MessageBeep(16);
          ShowMessage('Você já está no último registro do arquivo....');
        end
      Else
        begin
          CDS_Clientes.Last;

        end;

      Checa_Botoes;
      Painel_Registros();
    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.FormKeyPress(Sender: TObject; var Key: Char);
begin

  If (Key = #13) then Begin
    If Combobox_Clientes.Focused then Exit;
    if Combobox_Pesquisa.focused then Exit;

    Key := #0;
    Perform( Wm_NextDlgCtl, 0, 0 );
  end;


  If Key in [',','.'] then
    If ( Not(CDS_Clientes.state = dsInsert) And Not(CDS_Clientes.state = dsEdit) ) And
    Not(Combobox_Clientes.focused) and not(ComboBox_Pesquisa.Focused) and
    (PageControl1.ActivePageIndex = 0) then
      Key := FormatSettings.DecimalSeparator //DecimalSeparator
    Else
      Exit;


  If (Key = #27) And ( Not(CDS_Clientes.state = dsInsert) And Not(CDS_Clientes.state = dsEdit) ) And
  Not(Combobox_Clientes.focused) and Not(ComboBox_Pesquisa.Focused) then
    begin
     BitBtn_SairClick(Nil);

     Key := #0;
    end;

  //if Key = #27 then
  //  Close;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.Painel_Registros;
begin

  Panel_Registros.Caption := FormatFloat('0000', CDS_Clientes.RecNo)+'/'+FormatFloat('0000', CDS_Clientes.RecordCount);

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.BitBtn_BuscaCEPClick(Sender: TObject);
Var
  slDados : TStringList;
begin

  Try
    slDados := TStringList.Create;
    slDados := BuscarCEP_ViaCEP( Captura_SoNumeroString(CDS_Clientes.FieldByName('CEP').AsString) );

//    for i := 0 to slDados.Count-1 do
//        begin
//          showmessage('Dados: '+slDados[i] );
//        end;

    CDS_Clientes.FieldByName('Endereco').AsString   := slDados[1];
    CDS_Clientes.FieldByName('Logradouro').AsString := slDados[1];

    CDS_Clientes.FieldByName('Bairro').AsString     := slDados[3];
    CDS_Clientes.FieldByName('Cidade').AsString     := slDados[4];
    CDS_Clientes.FieldByName('Estado').AsString     := slDados[5];
    //CDS_Clientes.FieldByName('IBGE').AsString     := slDados[7]; // caso tivesse

    If (DBEdit_Numero.CanFocus) then
      DBEdit_Numero.SetFocus;

  Finally
    slDados.Destroy;
  End;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.BitBtn_SalvaXMLClick(Sender: TObject);
Var
  SaveDlgxml     : TSaveDialog;
begin
  SaveDlgXml := TSaveDialog.Create(FrmClientes);
  SaveDlgXml.FileName   :=  '';
  SaveDlgXml.Title      := 'Selecione o local aonde será salvo e informe o nome do.xml';
  SaveDlgXml.DefaultExt := '*.XML';
  SaveDlgXml.Filter     := 'Arquivos XML (*.XML)|*.XML|Arquivos TXT (*.TXT)|*.TXT';
  If not SaveDlgxml.Execute then
    Exit;

  //--

  Try
    Try
      If (SaveDlgxml.FileName = '') then
        begin
          MessageBeep(0);
          MessageDlg('Atenção...Selecione o local e informe o nome do .xml a salvar !', mtError, [mbOk], 0 );
          Exit;
        end;

      //--

      MessageBeep(16);
      If MessageDlg('Confirma a exportação dos dados para este arquivo .XML ?'+#13+#13+'Arquivo: '+SaveDlgxml.FileName+' ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
        begin
          Screen.Cursor := crHourGlass;

          CDS_Clientes.SaveToFile(SaveDlgxml.FileName);

          Screen.Cursor := crDefault;

          //--

          MessageBeep(0);
          MessageDlg('Dados exportados com sucesso !', mtInformation, [mbOK], 0);
        end;

    Except
      on ERRO : Exception do
        begin
          MessageBeep(0);
          MessageDlg('Atenção...'+#13+#13+'Não foi possivel importar este arquivo .XML !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);
        end;

    End;
  Finally
    SaveDlgxml.Free;
  End;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.BitBtn_CarregaXMLClick(Sender: TObject);
Var
  OpDgxml     : TOpenDialog;
  bApagaDados : Boolean;
begin
  OpDgxml := TOpenDialog.Create(FrmClientes);
  OpDgxml.FileName   :=  '';
  OpDgxml.Title      := 'Selecione o .xml';
  OpDgxml.DefaultExt := '*.XML';
  OpDgxml.Filter     := 'Arquivos XML (*.XML)|*.XML|Arquivos TXT (*.TXT)|*.TXT';
  If not OpDgxml.Execute then
    Exit;

  //--

  Try
    Try
      If (OpDgxml.FileName = '') then
        begin
          MessageBeep(0);
          MessageDlg('Atenção...Selecione o arquivo .xml a importar !', mtError, [mbOk], 0 );
          Exit;
        end;

      //--

      If ( CDS_Clientes.IsEmpty = False ) then
        begin
          MessageBeep(16);
          If MessageDlg('Deseja apagar os dados já cadastrados ?'+#13+#13+'Isto é perigoso pois podem haver registros de clientes com mesmo ID dos que serão importados !'+#13+#13+'Mantém estes dados assim mesmo ?' ,mtConfirmation,[mbYES,mbNO],0) = mrNo then
            bApagaDados := True;

        end;

      //--

      MessageBeep(16);
      If MessageDlg('Confirma a importação dos dados deste .XML ?'+#13+#13+'Arquivo: '+OpDgxml.FileName+' ?',mtConfirmation,[mbYES,mbNO],0) = mrYES then
        begin

          Screen.Cursor := crHourGlass;

          If bApagaDados then
            CDS_Clientes.EmptyDataSet;

          CDS_Clientes.LoadFromFile(OpDgxml.FileName);

          Screen.Cursor := crDefault;

          //--

          MessageBeep(0);
          MessageDlg('Dados importados com sucesso !', mtInformation, [mbOK], 0);

          Checa_Botoes;
          Painel_Registros();
        end;

    Except
      on ERRO : Exception do
        begin
          MessageBeep(0);
          MessageDlg('Atenção...'+#13+#13+'Não foi possivel importar este arquivo .XML !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);
        end;

    End;
  Finally
    OpDgxml.Free;
  End;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.BitBtn_LocalizarClick(Sender: TObject);
begin

  Bitbtn_Inclui.enabled        := False;
  Bitbtn_Altera.enabled        := False;
  Bitbtn_Salvar.enabled        := False;
  Bitbtn_Excluir.enabled       := False;
  BitBtn_Localizar.enabled     := False;
  Bitbtn_Sair.enabled          := False;
  DBNavigator_Clientes.enabled := False;

  ComboBox_Pesquisa.Enabled   := True;
  Combobox_Pesquisa.Setfocus;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.BitBtn_SairClick(Sender: TObject);
begin

  MessageBeep(16);
  If (MessageDlg('Deseja realmente sair da aplicação ?', mtConfirmation, [mbYes, mbNo], 0) = mrNo) then
    Exit;

  //--

  If CDS_Clientes.Active then CDS_Clientes.Close;

  Close;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.Checa_Botoes;
begin

  If CDS_Clientes.RecordCount = 0 then
    begin
      BitBtn_Inclui.enabled          := True;
      BitBtn_Altera.enabled          := False;
      BitBtn_Salvar.enabled          := False;
      BitBtn_Excluir.enabled         := False;
      BitBtn_Localizar.enabled       := False;
      BitBtn_Sair.enabled            := True;

      DBNavigator_Clientes.enabled   := False;

      ActAlterar.Enabled             := False;
      ActSalvar.Enabled              := False;
      ActExcluir.Enabled             := False;

      BitBtn_SalvaXML.enabled        := False;

      TabSheet_EnvioEmail.TabVisible := False;
    end
  Else
    begin
      BitBtn_Inclui.enabled          := True;
      BitBtn_Altera.enabled          := True;
      BitBtn_Salvar.enabled          := False;
      BitBtn_Excluir.enabled         := True;
      BitBtn_Localizar.enabled       := True;

      //--

      BitBtn_Sair.enabled            := True;
      DBNavigator_Clientes.enabled   := True;

      BitBtn_SalvaXML.enabled        := True;

      ActAlterar.Enabled             := True;
      ActSalvar.Enabled              := False;
      ActExcluir.Enabled             := True;

      TabSheet_EnvioEmail.TabVisible := True;
    end;

  BitBtn_CarregaXML.Enabled := True;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ComboBox_ClientesChange(Sender: TObject);
begin

  If (ComboBox_Pesquisa.ItemIndex = 0) then
    CDS_Clientes.Locate('NOME', ComboBox_Clientes.Text, [loCaseInsensitive, loPartialKey])

  Else If (ComboBox_Pesquisa.ItemIndex = 1) then
    CDS_Clientes.Locate('CPF', ComboBox_Clientes.Text, [loPartialKey])

  Else If (ComboBox_Pesquisa.ItemIndex = 2) then
    CDS_Clientes.Locate('RG', ComboBox_Clientes.Text, [loPartialKey])

  Else If (ComboBox_Pesquisa.ItemIndex = 3) then
    CDS_Clientes.Locate('EMAIL', ComboBox_Clientes.Text, [loCaseInsensitive, loPartialKey]);

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ComboBox_ClientesKeyPress(Sender: TObject;
  var Key: Char);
Var
  Prox         : Boolean;
  cTextoBuscar : string;
begin

  with (Sender as Tcombobox) do
    begin
      Autopreencher := True;
      Prox          := false;

      Case Ord(key) of
        Vk_return : begin
                      Selstart := length(text);
                      Sellength := 0;
                      Key := #0;
                      DroppedDown := false;
                      Prox := True;

                      cTextoBuscar := ComboBox_Clientes.Text;

                      if (ComboBox_Pesquisa.ItemIndex = 0) then
                        begin
                          If not( CDS_Clientes.Locate('NOME', Combobox_Clientes.Text, [loCaseInsensitive, loPartialKey]) ) then
                            begin
                              MessageBeep(0);
                              MessageDlg('Cliente com o Nome: '+Combobox_Clientes.Text+', não encontrado !', mtWarning, [mbOK], 0);

                              //--

                              MessageBeep(16);
                              If (MessageDlg('Deseja procurar algum outro mais aproximado ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes ) then
                                CDS_Clientes.Locate('NOME', Combobox_Clientes.Text, [loCaseInsensitive, loPartialKey]);

                            end

                        end

                      Else if (ComboBox_Pesquisa.ItemIndex = 1) then
                        begin
                          If not( CDS_Clientes.Locate('CPF', Combobox_Clientes.Text, [loPartialKey]) ) then
                            begin
                              MessageBeep(32);
                              MessageDlg('Cliente com o CPF: '+Combobox_Clientes.Text+', não encontrado !', mtWarning, [mbOK], 0);
                            end;

                        end

                      Else if (ComboBox_Pesquisa.ItemIndex = 2) then
                        begin
                          If not( CDS_Clientes.Locate('RG', Combobox_Clientes.Text, [loPartialKey]) ) then
                            begin
                              MessageBeep(32);
                              MessageDlg('Cliente com o RG: '+Combobox_Clientes.Text+', não encontrado !', mtWarning, [mbOK], 0);
                            end;

                        end

                      Else if (ComboBox_Pesquisa.ItemIndex = 3) then
                        begin
                          If not( CDS_Clientes.Locate('EMAIL', Combobox_Clientes.Text, [loPartialKey]) ) then
                            begin
                              MessageBeep(32);
                              MessageDlg('Cliente com o E-MAIL: '+Combobox_Clientes.Text+', não encontrado !', mtWarning, [mbOK], 0);
                            end;

                        end;

                      //--

                      ComboBox_Clientes.Text    := '';
                      ComboBox_Clientes.enabled := False;

                      ComboBox_Pesquisa.enabled := False;
                      Checa_Botoes;

                      Painel_Registros();

                      BitBtn_Inclui.SetFocus;
                      Exit;
                    end;

        Vk_back : Autopreencher := False;
        Vk_escape : begin
                      If Combobox_Clientes.Text <> '' then
                        Combobox_Clientes.Text := ''
                      Else
                        begin
                          Combobox_Clientes.enabled := False;
                          ComboBox_Pesquisa.enabled := False;

                          Checa_Botoes;
                          BitBtn_Inclui.Setfocus;
                        end;

                      Key := #0;
                    end;
      end;

      If (not Autopreencher) and ( Seltext <> '') then
        begin
          Text := Copy(Text,1,Selstart);
          Selstart := Length(text);
          Sellength := 0;
          Key := #0;
        end;
      end;

  If Prox then Findnextcontrol( Sender as TCombobox, True, True, False).SetFocus;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.ComboBox_PesquisaKeyPress(Sender: TObject;
  var Key: Char);
begin

  with (Sender as Tcombobox) do
    begin

      Case Ord(key) of
        Vk_return : begin
                      Combobox_Clientes.Text    := '';
                      Combobox_Clientes.enabled := True;

                      Combobox_Clientes.SetFocus;
                    end;

        Vk_back : Autopreencher := False;
        Vk_escape : begin
                      If Combobox_Clientes.Text <> '' then
                        Combobox_Clientes.Text := ''
                      Else
                        begin
                          Combobox_Pesquisa.enabled := False;
                          Combobox_Clientes.enabled  := False;

                          Checa_Botoes;
                          BitBtn_Inclui.Setfocus;

                        end;

                      Key := #0;
                    end;
      end;

      end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.DBEdit_NomeKeyPress(Sender: TObject; var Key: Char);
begin

  If Key = #27 then
    begin
      MessageBeep(16);
      If MessageDlg('Cancela esta Inclusão/Alteração ?', mtConfirmation, [mbYES,mbNO],0 ) = mrYES then
        begin
          CDS_Clientes.Cancel;
          //CDS_Clientes.Refresh;

          Habilita_Controles(false);
          BitBtn_Inclui.Setfocus;
        end;
      Key := #0;
      Exit;
    end;

  //--

  If (TDBEdit(Sender).Tag = 1) then
    begin

      If ( TDBEdit(Sender).Name = 'DBEdit_CPF') then
        begin
          // Valida só aceitando números...
          If Not(Key in ['0'..'9', #86, #118, #8, #13, #27 ]) then
            begin
              MessageBeep(0);
              MessageDlg('Tecla pressionada é inválida...É Permitido somente números !', mtError, [mbOk], 0 );

              Key := #0;
            end;

        end;


      // Valida só aceitando números...
      If Not(Key in ['0'..'9', #8, #13, #27 ]) then
        begin
          MessageBeep(0);
          MessageDlg('Tecla pressionada é inválida...É Permitido somente números !', mtError, [mbOk], 0 );

          Key := #0;
        end;

    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.DBGrid_ClientesTitleClick(Column: TColumn);
Var
  bReg_Bookmark : TBookmark;
begin

  If ( Column.FieldName <> '') then
    begin
      bReg_Bookmark := CDS_Clientes.GetBookmark;

      CDS_Clientes.IndexFieldNames := Column.FieldName;

      CDS_Clientes.GotoBookmark( bReg_Bookmark );
    end;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.DirectoryListBox1Change(Sender: TObject);
begin

  FileListBox1.Directory := DirectoryListBox1.Directory;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.DirectoryListBox1DblClick(Sender: TObject);
begin

  Edit_CaminhoSelecionado.Text := DirectoryListBox1.Directory;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.DriveComboBox1Change(Sender: TObject);
begin

  DirectoryListBox1.Drive := DriveComboBox1.Drive;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

function TFrmClientes.Valida_Campos : Boolean;
Var
  CDSBusca : TClientDataSet;
begin

  If DBEdit_Nome.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe o nome !!!', mtError, [mbOk], 0 );

      If DBEdit_Nome.Canfocus then
        DBEdit_Nome.SetFocus;

      Result := False;
      Exit;
    end
  Else
    begin
      CDSBusca := TClientDataSet.Create(nil);

      Try
        CDSBusca.CloneCursor(CDS_Clientes, false, false);

        If ( CDSBusca.Locate('Nome', DBEdit_Nome.Text, [loCaseInsensitive] ) ) and
        ( CDS_Clientes.FieldByName('ID').AsInteger <> CDSBusca.FieldByName('ID').AsInteger) then
          begin
            MessageBeep(0);
            Messagedlg('Nome já cadastrado !', mtError, [mbOk], 0 );

            If DBEdit_Nome.Canfocus then
              DBEdit_Nome.SetFocus;

            Result := False;
            Exit;
          end;

      Finally
        CDSBusca.Free;

      End;

    End;

  //--

  If (DBEdit_Identidade.Text = '') then
    begin
      MessageBeep(0);
      Messagedlg('Informe o RG !!!', mtError, [mbOk], 0 );

      If DBEdit_Identidade.Canfocus then
        DBEdit_Identidade.SetFocus;
      Result := False;
      Exit;
    end;


  If DBEdit_CPF.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe o CPF !!!', mtError, [mbOk], 0 );

      If DBEdit_CPF.Canfocus then
        DBEdit_CPF.SetFocus;
      Result := False;
      Exit;
    end
  Else
    begin

      If (DBEdit_CPF.Text = '   .   .   -  ') Or (Trim(DBEdit_CPF.Text) = '') then
        begin
          MessageBeep(0);
          Messagedlg('Atenção...'+#13+#13+'CPF deve ser informado...', mtError, [mbOk], 0 );

          If DBEdit_CPF.CanFocus then DBEdit_CPF.SetFocus;
          Result := False;
          Exit;
        end;

      // Valida o CPF para verificar se é válido...
      If ChecaCPF(DBEdit_CPF.Text, False) = False then
        begin
          MessageBeep(0);
          Messagedlg('Atenção...'+#13+#13+'CPF informado é inválido !', mtError, [mbOk], 0 );

          If DBEdit_CPF.CanFocus then DBEdit_CPF.SetFocus;
          Result := False;
          Exit;
        end;

      //--

      CDSBusca := TClientDataSet.Create(nil);
      Try
        CDSBusca.CloneCursor(CDS_Clientes, false, false);

        If ( CDSBusca.Locate('CPF', DBEdit_CPF.Text, [loCaseInsensitive] ) ) and
        ( CDS_Clientes.FieldByName('ID').AsInteger <> CDSBusca.FieldByName('ID').AsInteger) then
          begin
            MessageBeep(0);
            Messagedlg('CPF já cadastrado !', mtError, [mbOk], 0 );

            If DBEdit_Nome.Canfocus then
              DBEdit_Nome.SetFocus;

            Result := False;
            Exit;
          end;

      Finally
        CDSBusca.Free;

      End;

    end;


  If not(DBEdit_Email.ToString.IsEmpty) then
    begin
      If (ValidaEmail(DBEdit_Email.Text) = False ) then
        begin
          MessageBeep(0);
          Messagedlg('Informe um e-mail válido !', mtError, [mbOk], 0 );

          If DBEdit_Email.Canfocus then
            DBEdit_Email.SetFocus;
          Result := False;
          Exit;

        end;

    end;


  If DBEdit_Endereco.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe um endereço válido !', mtError, [mbOk], 0 );

      If DBEdit_Endereco.Canfocus then
        DBEdit_Endereco.SetFocus;
      Result := False;
      Exit;
    end;


  If (Captura_SoNumeroString(DBEdit_CEP.Text) = '') Or (Length(Captura_SoNumeroString(DBEdit_CEP.Text) ) <> 8) then
    begin
      MessageBeep(0);
      Messagedlg('Informe um CEP válido !', mtError, [mbOk], 0 );

      If DBEdit_CEP.Canfocus then
        DBEdit_CEP.SetFocus;
      Result := False;
      Exit;
    end;

  If DBEdit_Logradouro.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe um logradouro válido !', mtError, [mbOk], 0 );

      If DBEdit_Logradouro.Canfocus then
        DBEdit_Logradouro.SetFocus;
      Result := False;
      Exit;
    end;

  //--

  If DBEdit_Numero.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe o número do Logradouro !', mtError, [mbOk], 0 );

      If DBEdit_Numero.Canfocus then
        DBEdit_Numero.SetFocus;
      Result := False;
      Exit;
    end;

  //--

    If DBEdit_Bairro.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe o bairro !', mtError, [mbOk], 0 );

      If DBEdit_Bairro.Canfocus then
        DBEdit_Bairro.SetFocus;
      Result := False;
      Exit;
    end;

   //--

  If DBEdit_Cidade.Text = '' then
    begin
      MessageBeep(0);
      Messagedlg('Informe a cidade !', mtError, [mbOk], 0 );

      If DBEdit_Cidade.Canfocus then
        DBEdit_Cidade.SetFocus;
      Result := False;
      Exit;
    end;

  //--

  If DBComboBox_UF.ItemIndex = -1 then
    begin
      MessageBeep(0);
      Messagedlg('Informe a UF !', mtError, [mbOk], 0 );

      If DBComboBox_UF.Canfocus then
        DBComboBox_UF.SetFocus;
      Result := False;
      Exit;
    end;

  //--

  If DBComboBox_Pais.ItemIndex = -1 then
    begin
      MessageBeep(0);
      Messagedlg('Informe o País !', mtError, [mbOk], 0 );

      If DBComboBox_Pais.Canfocus then
        DBComboBox_Pais.SetFocus;
      Result := False;
      Exit;
    end;

  //--

  Result := True;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.Habilita_Controles( bHabilita : Boolean );
begin

  DBEdit_Nome.enabled            := bHabilita;
  DBEdit_Identidade.enabled      := bHabilita;
  DBEdit_CPF.enabled             := bHabilita;
  DBEdit_Telefone.enabled        := bHabilita;
  DBEdit_Email.enabled           := bHabilita;
  DBEdit_Endereco.enabled        := bHabilita;
  DBEdit_CEP.enabled             := bHabilita;
  BitBtn_BuscaCEP.enabled        := bHabilita;
  DBEdit_Logradouro.enabled      := bHabilita;
  DBEdit_Numero.enabled          := bHabilita;
  DBEdit_Complemento.enabled     := bHabilita;
  DBEdit_Bairro.enabled          := bHabilita;
  DBEdit_Cidade.enabled          := bHabilita;
  DBComboBox_UF.enabled          := bHabilita;
  DBComboBox_Pais.enabled        := bHabilita;

  TabSheet_EnvioEmail.TabVisible := not(bHabilita);

  if not(bHabilita) then
    Checa_Botoes;

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

function TFrmClientes.ChecaCPF(CPF : string; bAlerta : Boolean) : Boolean;
var
  TextCPF                   : string;
  f, Soma, Digito1, Digito2 : Integer;

begin
  Result := False;

  //--

  for f := 0 To 9 do
    begin
      if aVet_Sequencia[f] = CPF then
        begin
          Messagedlg('CPF inválido (nros. iguais) !!', mtError, [mbOk], 0 );
          Result := False;
        end;

    end;

   //--

  {verifica se existe carcteres inválidos}
  for f := 1 to Length(CPF) do
    if not (CPF[f] in ['0'..'9', '-', '.', ' ']) then
      Exit;
  {retira os caracteres não numéricos}
  TextCPF := '';
  for f := 1 to Length(CPF) do
    if CPF[f] in ['0'..'9'] then
      TextCPF := TextCPF + CPF[f];
  if TextCPF = '' then Result := True;
  if Length(TextCPF) <> 11 then Exit;
  {verifica primeiro digito}
  Soma := 0;
  for f := 1 to 9 do
    Soma := Soma + (StrToInt(TextCPF[f])*f);
  Digito1 := Soma mod 11;
  if Digito1 = 10 then Digito1 := 0;
  {verifica segundo digito}
  Soma := 0;
  for f := 1 to 8 do
    Soma := Soma + (StrToInt(TextCPF[f+1])*(f));
  Soma := Soma + (Digito1*9);
  Digito2 := Soma mod 11;
  if Digito2 = 10 then Digito2 := 0;
  {faz a validação}
  If (Digito1 = StrToInt(TextCPF[10])) and (Digito2 = StrToInt(TextCPF[11])) then
    Result := True
  Else
    begin
      If bAlerta then
        begin
          MessageBeep(0);
          Messagedlg('CPF inválido !!!!!'+#13+#13+'Atenção: Para esta seqüência de números os dois últimos dígitos'+#13+'corretos são: '+IntToStr(Digito1)+IntToStr(Digito2), mtError, [mbOk], 0 );
        end;
    end;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

function TFrmClientes.ValidaEmail(const s:string): Boolean;
Var
  lTemArroba : Boolean;
  Tamanho, i : integer;
begin

  lTemArroba := False;
  Tamanho := Length (Trim (S));

  //--

  For i := 0 to Tamanho do
    begin
       If Copy(S,i,1) = '@' then
         begin
           lTemArroba := True;
           Break;
         end;
    end;

  If lTemArroba then
    result := True
  Else
    Result := False;
end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

function TFrmClientes.Replicate( pString:String; xWidth:Integer ) :String;
Var
  nCount : Integer;
  pStr   : String;
begin
  pStr := '';
  For nCount := 1 To xWidth do pStr := pStr + pString;
  Result := pStr;
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.EntraemDigitacao( Sender : TObject );
begin

      If (Sender is TCombobox) then
        begin
          cCorFundo_ANT                    := (Sender as TCombobox).color;
          cCorLetra_ANT                    := (Sender as TCombobox).font.color;

          //--

          (Sender as TCombobox).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TCombobox).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TDBCombobox then
        begin
          cCorFundo_ANT                      := (Sender as TDBCombobox).color;
          cCorLetra_ANT                      := (Sender as TDBCombobox).font.color;

          //--

          (Sender as TDBCombobox).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TDBCombobox).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TDBEdit then
        begin
          cCorFundo_ANT                  := (Sender as TDBEdit).color;
          cCorLetra_ANT                  := (Sender as TDBEdit).font.color;

          //--

          (Sender as TDbEdit).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TDbEdit).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TDBLookupCombobox then
        begin
          cCorFundo_ANT                            := (Sender as TDBLookupComboBox).color;
          cCorLetra_ANT                            := (Sender as TDBLookupComboBox).font.color;

          //--

          (Sender as TDBLookUpCombobox).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TDBLookUpCombobox).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TDBMemo then
        begin
          cCorFundo_ANT                  := (Sender as TDBMemo).color;
          cCorLetra_ANT                  := (Sender as TDBMemo).font.color;

          //--

          (Sender as TDBMemo).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TDBMemo).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TEdit then
        begin
          cCorFundo_ANT                := (Sender as TEdit).color;
          cCorLetra_ANT                := (Sender as TEdit).font.color;

          //--

          (Sender as TEdit).color      := Entra_Digitacao_Cor_Fundo;
          (Sender as TEdit).font.color := Entra_Digitacao_Cor_Letra;
        end
      Else If Sender is TMaskEdit then
        begin
          cCorFundo_ANT                    := (Sender as TMaskEdit).color;
          cCorLetra_ANT                    := (Sender as TMaskEdit).font.color;

          //--

          (Sender as TMaskEdit).color      := clWindow;
          (Sender as TMaskEdit).font.color := Sai_Digitacao_Cor_Letra;
        end

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.SaidaDigitacao( Sender : TObject );
Var
  Cor_Fundo, Cor_Letra : TColor;
begin

  // Atribui a variável usada a cor salva ao usar a função EntraEmDigitacao()...
  Cor_Fundo := cCorFundo_ANT;
  Cor_Letra := cCorLetra_ANT;

  //---
  If (Sender is TCombobox) then
    begin
      (Sender as TCombobox).color      := Cor_Fundo;
      (Sender as TCombobox).font.color := Cor_Letra;
    end
  Else If Sender is TDBCombobox then
    begin
      (Sender as TDBCombobox).color      := Cor_Fundo;
      (Sender as TDBCombobox).font.color := Cor_Letra;
    end
  Else If Sender is TDBEdit then
    begin
      (Sender as TDbEdit).color      := Cor_Fundo;
      (Sender as TDbEdit).font.color := Cor_Letra;
    end
  Else If Sender is TDBLookupCombobox then
    begin
      (Sender as TDBLookUpCombobox).color      := Cor_Fundo;
      (Sender as TDBLookUpCombobox).font.color := Cor_Letra;
    end
  Else If Sender is TDBMemo then
    begin
      (Sender as TDBMemo).color      := Cor_Fundo;
      (Sender as TDBMemo).font.color := Cor_Letra;
    end
  Else If Sender is TEdit then
    begin
      (Sender as TEdit).color      := Cor_Fundo;
      (Sender as TEdit).font.color := Cor_Letra;
    end

end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.Spb_EnviaEmailClick(Sender: TObject);
Var
  cCaminhoArquivo, cCorpoEmail : String;
  cCPF, cFone                  : string;
begin

  If ( Trim(Edit_PortaSMTP.Text) = '') Or (Trim(Edit_ServidorSMTP.Text) = '')
  Or  (Trim(Edit_UsuarioSMTP.Text) = '') Or (Trim(Edit_SenhaSMTP.Text) = '') then
    begin
      MessageBeep(16);
      MessageDlg('Atenção...'+#13+#13+'Informe os dados da conta de e-mail para o envio !'+#13+#13+'Verifique...', mtWarning, [mbOK], 0);

      Exit;
    end;

  //--

  Try
    try
      Screen.Cursor := crHourGlass;

      cCaminhoArquivo := DirectoryListBox1.Directory;

      If (cCaminhoArquivo = '') Or ( DirectoryExists(cCaminhoArquivo) = False) then
        begin
          MessageBeep(16);
          MessageDlg('Atenção...'+#13+#13+'Caminho selecionado para gravação do .XML parece não existir ou não ter sido informado !'+#13+#13+'Verifique...', mtWarning, [mbOK], 0);
          Exit;
        end;

      //--

      MessageBeep(16);
      If (Messagedlg('Confirma a geração do arquivo e seu envio ao email do cliente ?', mtConfirmation, [mbYes, mbNo], 0) = mrNo ) then
        Exit;

      //--

      cCaminhoArquivo := cCaminhoArquivo + '\Cliente_'+FormatFloat('###########',CDS_Clientes.FieldByName('CPF').AsFloat)+'.xml';
      CDS_Clientes.SaveToFile( cCaminhoArquivo );

      cCorpoEmail := 'Seguem dados do: [ '+CDS_Clientes.FieldByName('Nome').AsString+' ] '#13+#13;


      cCPF := Copy(CDS_Clientes.FieldByName('CPF').AsString,1,3)+'.'+
              Copy(CDS_Clientes.FieldByName('CPF').AsString,4,3)+'.'+
              Copy(CDS_Clientes.FieldByName('CPF').AsString,7,3)+'-'+
              Copy(CDS_Clientes.FieldByName('CPF').AsString,10,2);

      cFone := '('+Copy(CDS_Clientes.FieldByName('Fone').AsString,01,02)+') '+
               Copy(CDS_Clientes.FieldByName('Fone').AsString,03,05)+'-'+
               Copy(CDS_Clientes.FieldByName('Fone').AsString,08,04);


      cCorpoEmail := cCorpoEmail+'Nome.................: '+CDS_Clientes.FieldByName('Nome').AsAnsiString+#13+
                                 'RG........................: '+CDS_Clientes.FieldByName('RG').AsAnsiString+#13+
                                 'CPF......................: '+cCPF+#13+ //CDS_Clientes.FieldByName('CPF').AsAnsiString+#13+
                                 'Telefone............: '+cFone+#13+ //CDS_Clientes.FieldByName('Fone').AsAnsiString+#13+
                                 'E-mail.................: '+CDS_Clientes.FieldByName('Email').AsAnsiString+#13+
                                 'Endereço..........: '+CDS_Clientes.FieldByName('Endereco').AsAnsiString+#13+
                                 'CEP......................: '+CDS_Clientes.FieldByName('CEP').AsAnsiString+#13+
                                 'Logradouro......: '+CDS_Clientes.FieldByName('Logradouro').AsAnsiString+#13+
                                 'Numero.............: '+CDS_Clientes.FieldByName('Numero').AsAnsiString+#13+
                                 'Complemento: '+CDS_Clientes.FieldByName('Complemento').AsAnsiString+#13+
                                 'Bairro.................: '+CDS_Clientes.FieldByName('Bairro').AsAnsiString+#13+
                                 'Cidade...............: '+CDS_Clientes.FieldByName('Cidade').AsAnsiString+#13+
                                 'UF........................: '+CDS_Clientes.FieldByName('Estado').AsAnsiString+#13+
                                 'País.....................: '+CDS_Clientes.FieldByName('Pais').AsAnsiString+#13+#13+#13+
                                 'Sem mais, ' + #13+#13 + 'InfoSistemas.'+#13+#13+#13+#13;

      Envia_Email( Trim(Edit_PortaSMTP.Text),
                   Trim(Edit_ServidorSMTP.Text),
                   Trim(Edit_UsuarioSMTP.Text),
                   Trim(Edit_SenhaSMTP.Text),
                   'meu_email_de_origem', //CDS_Clientes.FieldByName('email').AsString,
                   'Arquivo XML - '+CDS_Clientes.FieldByName('nome').AsString,
                   CDS_Clientes.FieldByName('email').AsString,
                   'Envio de arquivo .xml - Dados cadastrais',
                   cCorpoEmail+'Informação a InfoSistemas: Quais metas precisamos bater em nos próximos 3 meses ?',
                   cCaminhoArquivo, 'S');

      //--

      MessageBeep(0);
      MessageDlg('Geração de arquivo .XML e envio de e-mail efetuado com êxito !'+#13+#13+'O e-mail foi enviado para: '+CDS_Clientes.FieldByName('Email').AsString, mtInformation, [mbOK], 0);

      //--

      MessageBeep(0);
      if (Messagedlg('Deseja apagar o arquivo gerado no disco ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
        begin
          If FileExists(cCaminhoArquivo) then
            DeleteFile(cCaminhoArquivo);
        end;

    Except
      on ERRO : Exception do
        begin
          MessageBeep(0);
          MessageDlg('Atenção...'+#13+#13+'Não foi possível salvar o arquivo .XML e enviá-lo via e-mail !'+#13+#13+'Erro: '+Erro.Message, mtError, [mbOK], 0);
        end;
    End;
  Finally
    Screen.Cursor := crDefault;

    PageControl1.ActivePageIndex := 0;
    BitBtn_Inclui.SetFocus;
  End;
end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

procedure TFrmClientes.TabSheet_EnvioEmailEnter(Sender: TObject);
begin

  DirectoryListBox1.Directory := GetCurrentDir();
  Edit_CaminhoSelecionado.Text := DirectoryListBox1.Directory;

  if (Edit_PortaSMTP.CanFocus) then Edit_PortaSMTP.SetFocus;
  
end;

//***---****---***---****---***---****---***---****---***---****---***---****-//

procedure TFrmClientes.Proced_EntraemDigitacao(Sender: TObject);
begin

  EntraemDigitacao( Sender );

end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

procedure TFrmClientes.Proced_SaidaDigitacao(Sender: TObject);
begin

  SaidaDigitacao( Sender );

end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

function TFrmClientes.Captura_SoNumeroString( cTexto: string ) : string;
Var
  cCaracter  : string[01];
  cResultado : string;
  I          : integer;
  lRet       : Boolean;
begin

  lRet := True;

  For I := 1 To Length(cTexto) do
    begin
      cCaracter := UpperCase(Copy(cTexto,I,01));

      If (Ord(cTexto[i]) >= 48) And (Ord(cTexto[i]) <= 57)   then
        cResultado := cResultado+cCaracter;

    end;

  If Trim(cResultado) <> '' then
    Result := cResultado
  Else
    Result := '';

end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

function TFrmClientes.BuscarCEP_ViaCEP(UmCEP: string): TStringList;
var
  data          : TJSONObject;
  RESTClient1   : TRESTClient;
  RESTRequest1  : TRESTRequest;
  RESTResponse1 : TRESTResponse;
  Endereco      : TStringList;
begin
  RESTClient1           := TRESTClient.Create(nil);
  RESTRequest1          := TRESTRequest.Create(nil);
  RESTResponse1         := TRESTResponse.Create(nil);
  RESTRequest1.Client   := RESTClient1;
  RESTRequest1.Response := RESTResponse1;

  RESTClient1.BaseURL   := 'viacep.com.br/ws/' + UmCEP + '/json/';

  RESTRequest1.Execute;
  data := RESTResponse1.JSONValue as TJSONObject;

  Try
    Endereco := TStringList.Create;
    If Assigned(data) then
      begin

        Try
          Endereco.Add(data.Values['cep'].Value);
        Except
          on Exception do
            Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['logradouro'].Value);
        Except
          on Exception do
            Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['complemento'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['bairro'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['localidade'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['uf'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['unidade'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['ibge'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

        Try
          Endereco.Add(data.Values['gia'].Value);
        except
         on Exception do
           Endereco.Add('');
        end;

      end;
  finally
    FreeAndNil(data);
  end;
  Result := Endereco;

end;

//***---***---***---***---***---***---***---***---***---***---***---***---***-//

end.
