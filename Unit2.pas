unit Unit2;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompButton,
  IWCompListbox, IWCompEdit, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWCompLabel, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, IWGrids, ComObj, Dialogs,
  Variants, Math, IWExtCtrls, jpeg, windows, ActiveX,
  IWWebGrid, IWAdvWebGrid, IWCompTabControl, IWAdvWebGridExcel, IWCompText,
  IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent, IWBaseHTML40Component,
  chart, series, pngimage,
  IWCompProgressBar, ExtCtrls, IWHTMLControls, TeeGDIPlus, TeEngine, TeeProcs,IWTMSPopup, IWHelpTip, Menus,
  IWCompMemo;

type
  TIWForm2 = class(TIWAppForm)
    rgnprob: TIWRegion;
    lblWarning: TIWLabel;
    btnInputManual: TIWButton;
    btnImportar: TIWButton;
    IWImage9: TIWImage;
    IWLabel70: TIWLabel;
    IWImage10: TIWImage;
    IWImage27: TIWImage;
    IWLabel12: TIWLabel;
    IWRegion16: TIWRegion;
    IWRegion3: TIWRegion;
    IWText2: TIWText;
    rgninput: TIWRegion;
    IWImage5: TIWImage;
    IWLabel7: TIWLabel;
    IWLink2: TIWLink;
    IWRegion4: TIWRegion;
    rgrpPeriodo: TIWRadioGroup;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    IWLabel3: TIWLabel;
    IWLabel4: TIWLabel;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWLabel9: TIWLabel;
    edtprodregular: TIWEdit;
    edthorasextras: TIWEdit;
    edtsubcontrat: TIWEdit;
    edtestoq: TIWEdit;
    edtfunc: TIWEdit;
    btnavancaretapa1: TIWButton;
    btnredefiniretapa1: TIWButton;
    btnsalvaretapa1: TIWButton;
    IWLink5: TIWLink;
    IWLink6: TIWLink;
    IWLabel14: TIWLabel;
    IWLabel15: TIWLabel;
    edtinifunc: TIWEdit;
    rgndemanda: TIWRegion;
    IWImage6: TIWImage;
    IWLabel16: TIWLabel;
    IWLink7: TIWLink;
    IWLink9: TIWLink;
    IWLabel25: TIWLabel;
    IWRegion2: TIWRegion;
    grddemanda: TTIWAdvWebGrid;
    btnsalvaretapa2: TIWButton;
    btnredefiniretapa2: TIWButton;
    imgdemanda: TIWImage;
    btnavancaretapa2: TIWButton;
    IWRegion20: TIWRegion;
    txtOPRSA: TIWText;
    txtCPRSA: TIWText;
    IWImage8: TIWImage;
    IWImage24: TIWImage;
    lblLRSA: TIWLabel;
    IWImage11: TIWImage;
    IWText1: TIWText;
    rgnfuncaoobjetivo: TIWRegion;
    IWImage14: TIWImage;
    IWLabel17: TIWLabel;
    IWLink8: TIWLink;
    IWLink10: TIWLink;
    IWLabel18: TIWLabel;
    btnsalvaretapa3: TIWButton;
    IWRegion5: TIWRegion;
    cmbbxObjetivos: TIWComboBox;
    IWLabel19: TIWLabel;
    grdfuncaoobjetivo: TTIWAdvWebGrid;
    btnreiniciaretapa3: TIWButton;
    btnexecutar: TIWButton;
    IWLabel20: TIWLabel;
    chrtDemanda: TChart;
    LineSeries2: TLineSeries;
    LineSeries3: TLineSeries;
    Series1: TLineSeries;
    cbxproducaoregular: TIWCheckBox;
    cbxHorasExtras: TIWCheckBox;
    cbxSubcontratacao: TIWCheckBox;
    cbxEstoque: TIWCheckBox;
    lstbxFuncoesObjetivo: TIWListbox;
    cbxcontrat: TIWCheckBox;
    cbxdem: TIWCheckBox;
    IWLabel8: TIWLabel;
    edtnomeproblema: TIWEdit;
    rgnintro: TIWRegion;
    IWLabel21: TIWLabel;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    txtwarnning: TIWText;
    IWButton3: TIWButton;
    IWLabel46: TIWLabel;
    IWImage7: TIWImage;
    IWLabel52: TIWLabel;
    IWImage16: TIWImage;
    IWImage17: TIWImage;
    iwrgnCaixaDialogoRU: TIWRegion;
    lblWg: TIWLabel;
    IWRegionru: TIWRegion;
    txtCaixaDialogoRU: TIWText;
    btnBackRU: TIWButton;
    btncontinueRU: TIWButton;
    IWRegion9: TIWRegion;
    IWRegion6: TIWRegion;
    rgnlogin: TIWRegion;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    btnEnter: TIWButton;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    lblwarningpassword: TIWLabel;
    lblwarningemail: TIWLabel;
    IWButton4: TIWButton;
    txtpassrecover: TIWText;
    txtnewcode: TIWText;
    lnkesqueceusenha: TIWLink;
    IWRegion19: TIWRegion;
    iwrgnCaixaDialogoRU2: TIWRegion;
    lblWg2: TIWLabel;
    IWRegionru2: TIWRegion;
    txtCaixaDialogoRU2: TIWText;
    btnBackRU2: TIWButton;
    btncontinueRU2: TIWButton;
    IWImage18: TIWImage;
    rgnretomada: TIWRegion;
    lblChooseproj: TIWLabel;
    lstbxProjects: TIWListbox;
    btnReset3: TIWButton;
    btnLogoutRgnLogin: TIWButton;
    IWLabel26: TIWLabel;
    lblMexplanation: TIWLabel;
    btnauxchoose: TIWButton;
    IWRegion1: TIWRegion;
    IWRegion8: TIWRegion;
    IWRegion23: TIWRegion;
    IWImage20: TIWImage;
    IWLink4: TIWLink;
    grdmatrizconseq: TTIWAdvWebGrid;
    grdplanos: TTIWAdvWebGrid;
    lnkPlanos: TIWLink;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    TIWAdvWebGridExcelIO2: TTIWAdvWebGridExcelIO;
    IWTimer1: TIWTimer;
    SaveDialog1: TFileSaveDialog;
    IWImage23: TIWImage;
    rgnOpcaoInput: TIWRegion;
    lblpleasechoose: TIWLabel;
    btnImport: TIWButton;
    btnRegisternew: TIWButton;
    btnContinueregist: TIWButton;
    btnLogoutRgnOpcaoInput: TIWButton;
    IWImage37: TIWImage;
    IWImage38: TIWImage;
    IWLabel22: TIWLabel;
    IWImage46: TIWImage;
    IWImage47: TIWImage;
    IWImage48: TIWImage;
    btnlogoutescolha: TIWButton;
    IWRegion7: TIWRegion;
    IWRegion22: TIWRegion;
    IWRegion29: TIWRegion;
    IWImage21: TIWImage;
    IWRegion10: TIWRegion;
    IWRegion11: TIWRegion;
    IWRegion14: TIWRegion;
    rgnnovoproblema: TIWRegion;
    IWImage28: TIWImage;
    IWLink11: TIWLink;
    IWLabel33: TIWLabel;
    IWLink13: TIWLink;
    btnSaveProblem: TIWButton;
    btnredefinirnovoproblema: TIWButton;
    IWLink14: TIWLink;
    IWLabel34: TIWLabel;
    lstbxperformanceobj: TIWListbox;
    cmbbxperformanceobj: TIWComboBox;
    btnobj: TIWButton;
    IWLabel36: TIWLabel;
    edtDecisionProblem: TIWEdit;
    IWRegion17: TIWRegion;
    rgnotherperformanceobj: TIWRegion;
    IWLabel37: TIWLabel;
    IWLabel38: TIWLabel;
    edtotherperformanceobj: TIWEdit;
    IWLink15: TIWLink;
    btnperformanceobj: TIWButton;
    IWRegion21: TIWRegion;
    IWLabel23: TIWLabel;
    memoContext: TIWMemo;
    IWLabel24: TIWLabel;
    rgnactorroles: TIWRegion;
    grdactorroles: TTIWAdvWebGrid;
    IWLabel27: TIWLabel;
    cmbxactor: TIWComboBox;
    IWLabel28: TIWLabel;
    edtnewactor: TIWEdit;
    btnNewActorRole: TIWButton;
    btnActorDescription: TIWButton;
    IWLabel29: TIWLabel;
    edtActorDescription: TIWEdit;
    IWImage15: TIWImage;
    IWLink1: TIWLink;
    btnexportar: TIWButton;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWRegion13: TIWRegion;
    IWImage12: TIWImage;
    IWImage13: TIWImage;
    IWImage19: TIWImage;
    IWImage22: TIWImage;
    IWImage3: TIWImage;
    IWImage4: TIWImage;
    IWRegion12: TIWRegion;
    IWImage29: TIWImage;
    IWImage30: TIWImage;
    IWImage25: TIWImage;
    IWImage26: TIWImage;
    procedure rgninputCreate(Sender: TObject);
    procedure rgnprobCreate(Sender: TObject);
    procedure btnInputManualClick(Sender: TObject);
    procedure IWLink2Click(Sender: TObject);
    procedure edtprodregularAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edthorasextrasAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edtsubcontratAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edtestoqAsyncKeyPress(Sender: TObject; EventParams: TStringList);
    procedure edtfuncAsyncKeyPress(Sender: TObject; EventParams: TStringList);
    procedure edtinifuncAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure rgndemandaCreate(Sender: TObject);
    procedure rgnfuncaoobjetivoCreate(Sender: TObject);
    procedure btnsalvaretapa1Click(Sender: TObject);
    procedure rgrpPeriodoClick(Sender: TObject);
    procedure btnredefiniretapa1Click(Sender: TObject);
    procedure btnavancaretapa1Click(Sender: TObject);
    procedure cbxproducaoregularClick(Sender: TObject);
    procedure cbxHorasExtrasClick(Sender: TObject);
    procedure cbxSubcontratacaoClick(Sender: TObject);
    procedure cbxEstoqueClick(Sender: TObject);
    procedure IWButton1Click(Sender: TObject);
    procedure btnsalvaretapa2Click(Sender: TObject);
    procedure btnredefiniretapa2Click(Sender: TObject);
    procedure btnavancaretapa2Click(Sender: TObject);
    procedure btnsalvaretapa3Click(Sender: TObject);
    procedure btnredefiniretapa3Click(Sender: TObject);
    procedure btnreiniciaretapa3Click(Sender: TObject);
    procedure btnexecutarClick(Sender: TObject);
    procedure Exportar;
    procedure SalvarBD;
    procedure cbxcontratClick(Sender: TObject);
    procedure cbxdemClick(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure rgnintroCreate(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure btnRegisterClick(Sender: TObject);
    procedure btncontinueRUClick(Sender: TObject);
    procedure btnBackRUClick(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure rgnloginCreate(Sender: TObject);
    procedure lnkesqueceusenhaClick(Sender: TObject);
    procedure btnBackRU2Click(Sender: TObject);
    procedure btncontinueRU2Click(Sender: TObject);
    procedure IWButton4Click(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure rgnretomadaCreate(Sender: TObject);
    procedure IWLink5Click(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure btnauxchooseClick(Sender: TObject);
    procedure IWLink6Click(Sender: TObject);
    procedure lnkplanosClick(Sender: TObject);
    procedure IWTimer1Timer(Sender: TObject);
    function TrocaVP(Valor: string): String;
    function TrocaPV(Valor: string): String;
    procedure LimpaVariaveis;
    procedure IWAppFormCreate(Sender: TObject);
    procedure btnRegisternewClick(Sender: TObject);
    procedure btnContinueregistClick(Sender: TObject);
    procedure rgnOpcaoInputCreate(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure btnobjClick(Sender: TObject);
    procedure btnperformanceobjClick(Sender: TObject);
    procedure IWLink15Click(Sender: TObject);
    procedure IWLink13Click(Sender: TObject);
    procedure btnNewActorRoleClick(Sender: TObject);
    procedure btnActorDescriptionClick(Sender: TObject);
    procedure rgnnovoproblemaCreate(Sender: TObject);
    procedure edtnewactorAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure IWLink11Click(Sender: TObject);
    procedure btnSaveProblemClick(Sender: TObject);
    procedure btnredefinirnovoproblemaClick(Sender: TObject);
    procedure btnexportarClick(Sender: TObject);

  public

  end;

implementation

uses UserSessionUnit, ServerController, untResultados;

{$R *.dfm}

function TIWForm2.TrocaVP(Valor: string): String;
var
  i: integer;
begin
  if Valor <> '' then
  begin
    for i := 0 to Length(Valor) do
    begin
      if Valor[i] = ',' then
      begin
        Valor[i] := '.';
      end
      else if Valor[i] = ' .' then
      begin
        Valor[i] := '.';
      end;
    end;
  end;
  Result := Valor;
end;

function TIWForm2.trocapv(Valor: string): String;
var
  i: integer;
begin
  if Valor <> '' then
  begin
    for i := 0 to Length(Valor) do
    begin
      if (Valor[i] = '.') or (Valor[i] = ' ,')  then
      begin
        Valor[i] := ',';
      end // if
    end; // i
  end; // if valor
  Result := Valor;
end;


procedure TIWForm2.btnActorDescriptionClick(Sender: TObject);
begin
if cmbxactor.ItemIndex <> -1 then
begin
grdactorroles.RowCount := grdactorroles.RowCount + 1;
grdactorroles.TotalRows := grdactorroles.RowCount + 1;
SetLength(TIWUserSession(WebApplication.Data).Actor, grdactorroles.RowCount);
SetLength(TIWUserSession(WebApplication.Data).ActorDescription, grdactorroles.RowCount);
TIWUserSession(WebApplication.Data).Actor[grdactorroles.RowCount-1] := cmbxactor.Text;
TIWUserSession(WebApplication.Data).ActorDescription[grdactorroles.RowCount-1] := edtActorDescription.Text;
grdactorroles.Cells[0,grdactorroles.RowCount-1] := cmbxactor.Text;
grdactorroles.Cells[1, grdactorroles.RowCount-1] := edtActorDescription.Text;
edtActorDescription.Text := '';
edtnewactor.Text  := '[In case the actor role is not included in the list]';
end;
end;

procedure TIWForm2.btnauxchooseClick(Sender: TObject);
var
i,j : integer;
Id_Problema : string;
begin
With TIWUserSession(WebApplication.Data) do
begin
with UserSession.ZQuery1 do
  begin
   SetLength(Limites, 5);
   Setlength(Custos, 6);
   Id_Problem := idproblemasuser[lstbxProjects.ItemIndex];

  SQL.Text := 'SELECT * FROM restrictions WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        + '"';
      ExecSQL;
      open;
      Periodo := strtoint(FieldByName('Total_Period').AsString);
      Limites[0] := strtofloat(FieldByName('Regular_Production').AsString);
      Limites[1] := strtofloat(FieldByName('Overtime_Production').AsString);
      Limites[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
      Limites[3] := strtofloat(FieldByName('Stock').AsString);
      Limites[4] := strtofloat(FieldByName('Max_Employees').AsString);
      NinicialFunc := strtoint(FieldByName('Ini_Employees').AsString);
      close;


    SetLength(Demanda, Periodo);
    for i := 0 to Periodo - 1 do
    begin
    SQL.Text := 'SELECT * FROM demand WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        + '" AND Period ="' + inttostr(i+1) + '"';
      ExecSQL;
      open;
      Demanda[i] := strtofloat(FieldByName('Demand').AsString);
      close;
    end;






      SetLength(Custos, 6);


      SQL.Text := 'SELECT * FROM obj_functions WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        +'"';
      ExecSQL;
      open;
      First;
      i := 0;
      j := 0;
      while not eof do
      begin
      SetLength(ObjSelecionados, i+1);
      TIWUserSession(WebApplication.Data).ObjSelecionados[i] := FieldByName('Objective').AsString;
      if ObjSelecionados[i] = 'Cost' then
      begin
        TIWUserSession(WebApplication.Data).Custos[0] := strtofloat(FieldByName('Regular_Production').AsString);
        Custos[1] := strtofloat(FieldByName('Overtime_Production').AsString);
        Custos[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
        Custos[3] := strtofloat(FieldByName('Stock').AsString);
        Custos[4] := strtofloat(FieldByName('Hiring').AsString);
        Custos[5] := strtofloat(FieldByName('Resignation').AsString);
      end
      else
      begin
       SetLength(ParametrosFuncaoObjetivo, j+1, 6);
       TIWUserSession(WebApplication.Data).ParametrosFuncaoObjetivo[j,0] := strtofloat(FieldByName('Regular_Production').AsString);
       ParametrosFuncaoObjetivo[j,1] := strtofloat(FieldByName('Overtime_Production').AsString);
       ParametrosFuncaoObjetivo[j,2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
       ParametrosFuncaoObjetivo[j,3] := strtofloat(FieldByName('Stock').AsString);
       ParametrosFuncaoObjetivo[j,4] := strtofloat(FieldByName('Hiring').AsString);
       ParametrosFuncaoObjetivo[j,5] := strtofloat(FieldByName('Resignation').AsString);
       inc(j);
      end;
      inc(i);
      Next;
      end;
      Close;

      rgninput.Visible := True;
      rgnretomada.Visible := False;

      if Periodo = 12 then
      rgrpPeriodo.ItemIndex := 0;
      if Periodo = 6 then
      rgrpPeriodo.ItemIndex := 1;
      if Periodo = 4 then
      rgrpPeriodo.ItemIndex := 2;

      edtnomeproblema.Text := nameproblem[lstbxProjects.ItemIndex];

      if (Limites[0] = 0) or (Custos[0] = 0) then
      begin
        cbxproducaoregular.Checked := False;
        cbxproducaoregularClick(Self);
      end;

      if (Limites[1] = 0) or (Custos[1] = 0) then
      begin
        cbxHorasExtras.Checked := False;
        cbxHorasExtrasClick(Self);
      end;

      if (Limites[2] = 0) or (Custos[2] = 0) then
      begin
        cbxSubcontratacao.Checked := False;
        cbxSubcontratacaoClick(Self);
      end;

       if (Limites[3] = 0) or (Custos[3] = 0) then
      begin
        cbxEstoque.Checked := False;
        cbxEstoqueClick(Self);
      end;

      if (Custos[4] = 0) then
      begin
        cbxcontrat.Checked := False;
        cbxcontratClick(Self);
      end;

      if (Custos[5] = 0) then
      begin
        cbxdem.Checked := False;
        cbxdemclick(Self);
      end;


      edtprodregular.Text := FloatToStr(limites[0]);
      edthorasextras.Text := FloatToStr(limites[1]);
      edtsubcontrat.Text := FloatToStr(limites[2]);
      edtestoq.Text := FloatToStr(limites[3]);
      edtfunc.Text := FloatToStr(limites[4]);
      edtinifunc.Text := IntToStr(NinicialFunc);
      edtprodregular.Text := trocavp(edtprodregular.Text);
      edthorasextras.Text := trocavp(edthorasextras.Text);
      edtsubcontrat.Text := trocavp(edtsubcontrat.Text);
      edtestoq.Text := trocavp(edtestoq.Text);
      edtfunc.Text := trocavp(edtfunc.Text);
      edtinifunc.Text := trocavp(edtinifunc.Text);
      btnsalvaretapa1Click(Self);

      lstbxperformanceobj.Items.Clear;
      lstbxperformanceobj.Items.Add('Selected Objectives:');
      for j := 0 to length(ObjSelecionados) - 1 do
      begin
        lstbxperformanceobj.Items.Add(ObjSelecionados[j]);
        for i := 0 to cmbbxperformanceobj.Items.Count - 1 do
        begin
          cmbbxperformanceobj.ItemIndex := i;
          if cmbbxperformanceobj.Text = objSelecionados[j] then
          begin
            cmbbxperformanceobj.Items.Delete(cmbbxperformanceobj.ItemIndex);
          end;
        end;
      end;
      btnavancaretapa1.Enabled := True;
      datainicio := datetostr(now);
  end;

  end;
end;

procedure TIWForm2.LimpaVariaveis;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Finalize(Limites);
    Finalize(Demanda);
    Finalize(ObjSelecionados);
    Finalize(OrdemSel);
    Finalize(Populacao);
    Finalize(Population);
    Finalize(PerformancePopulacao);
    Finalize(PopulationPerformance);
    Finalize(NovaGeracao);
    Finalize(MatrizFinal);
    Finalize(Custos);
    Finalize(ParametrosFuncaoObjetivo);
    Finalize(PerformanceNormalizada);
    Finalize(PopulationRanking);
    Finalize(OrdemPopulacao);
    Finalize(PopulationDistance);
    Finalize(DistanciaPopulacional);
    Finalize(nameproblem);
    Finalize(idproblemasuser);
    Finalize(datainicioproblema);
    Finalize(datafinalproblema);
    Finalize(VetorGrafico);
    Finalize(indexROC);
    Finalize(maior);
    Finalize(menor);
    Finalize(valor_razao);
    Finalize(BConseqOrd);
    Finalize(WConseqOrd);
    Finalize(peso);


  end;
end;

procedure TIWForm2.btnavancaretapa1Click(Sender: TObject);
var
i : integer;
number : double;
begin
With TIWUserSession(WebApplication.Data) do
  begin
  edtprodregular.Text := trocavp(edtprodregular.Text);
  edthorasextras.Text := trocavp(edthorasextras.Text);
  edtsubcontrat.Text := trocavp(edtsubcontrat.Text);
  edtestoq.Text := trocavp(edtestoq.Text);
  edtfunc.Text := trocavp(edtfunc.Text);
  edtinifunc.Text := trocavp(edtinifunc.Text);
  if (TryStrToFloat(edtprodregular.Text, number) = True) and
   (TryStrToFloat(edthorasextras.Text, number) = True) and
   (TryStrToFloat(edtsubcontrat.Text, number) = True) and
   (TryStrToFloat(edtestoq.Text, number) = True) and
   (TryStrToFloat(edtfunc.Text, number) = True) and
   (TryStrToFloat(edtinifunc.Text, number) = True) then
  begin
  Setlength(Limites, 5);

  Limites[0] := Strtofloat(edtprodregular.Text); // ** Max prod regular
  Limites[1] := Strtofloat(edthorasextras.Text); // ** Max prod h extra
  Limites[2] := Strtofloat(edtsubcontrat.Text); // ** Maxima prod subcontratada
  Limites[3] := Strtofloat(edtestoq.Text); // **Máximo estoque por período
  Limites[4] := Strtofloat(edtfunc.Text); // ** Número máximo de funcionários
  NinicialFunc := Strtoint(edtinifunc.Text); // ** Número inicial de funcionários
  edtprodregular.Enabled := False;
  edthorasextras.Enabled := False;
  edtsubcontrat.Enabled := False;
  edtestoq.Enabled := False;
  edtfunc.Enabled := False;
  edtinifunc.Enabled := False;

  btnsalvaretapa1.enabled := False;
  //if (lstbxobjetivos.Items.Count > 2) and (rgrpPeriodo.ItemIndex <> -1) then
  //btnavancaretapa1.Enabled := True;
  end
  else
  begin
  WebApplication.ShowMessage('Fill in the problem restrictions correctly!');
  exit;
  end;
  end;

 if rgrpPeriodo.ItemIndex = -1 then
 begin
  WebApplication.ShowMessage('Select the period to continue!');
  exit;
  end;

  

  if (edtnomeproblema.Text = '') then
  begin
  WebApplication.ShowMessage('Fill in the problem name to continue!');
  exit;
  end;

With TIWUserSession(WebApplication.Data) do
begin
rgninput.Visible := False;
rgndemanda.Visible := True;
grddemanda.TotalRows := Periodo;
grddemanda.rowcount := Periodo;
Problema := edtnomeproblema.Text;
Chrtdemanda.Series[0].Clear;
Chrtdemanda.Series[1].Clear;
Chrtdemanda.Series[2].Clear;
Chrtdemanda.Series[0].Visible := False;
Chrtdemanda.Series[1].Visible := False;
Chrtdemanda.Series[2].Visible := True;
for i := 0 to Periodo -1  do
begin
chrtDemanda.Series[2].AddXY(i+1, Limites[0]*limites[4] + Limites[1]*limites[4] + Limites[2]);
grddemanda.Cells[0,i] := inttostr(i+1);
end;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));
SetLength(ParametrosFuncaoObjetivo, Length(ObjSelecionados)-1, 6);
cmbbxObjetivos.Items.Clear;
for I := 0 to Length(ObjSelecionados)-1 do
cmbbxobjetivos.Items.Add(objselecionados[i]);
cmbbxobjetivos.ItemIndex := 0;

if (retomada = True) or (problemapadrao = true) then
begin
  for i := 0 to Periodo - 1 do
  grddemanda.DynEdits[1,i] := FloatToStr(Demanda[i]);
  grddemanda.DynEdits[1,i] := trocavp(grddemanda.DynEdits[1,i]);
  btnsalvaretapa2click(self);
end;

end;
end;

procedure TIWForm2.btnavancaretapa2Click(Sender: TObject);
var
i, aux : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := True;
 SetLength(OrdemSel, cmbbxObjetivos.Items.Count);
if (retomada = True) or (problemapadrao = true)  then
begin
  for i := 0 to cmbbxObjetivos.Items.Count - 1 do
  begin
    ordemsel[i] := i;
    cmbbxObjetivos.ItemIndex := i;
    if i = 0 then
    begin
      lstbxFuncoesObjetivo.Items.Add('Min. ' + cmbbxObjetivos.Text + ' = ' + floattostr(custos[0]) +
      '*x1 + '+ floattostr(custos[1]) +'*x2 + '+ floattostr(custos[2])+ '*x3 + '+ floattostr(custos[3])+ '*x4 + '
      + floattostr(custos[4])+ '*x5 + ' + floattostr(custos[5])+ '*x6');
    end
    else
    begin
    aux := i;
      lstbxFuncoesObjetivo.Items.Add('Max. ' + cmbbxObjetivos.Text + ' = ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 0]) +
      '*x1 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 1]) +'*x2 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 2])+ '*x3 '
      + floattostr(ParametrosFuncaoObjetivo[aux-1, 3]) +'*x4 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 4]) +'*x5 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 5]) +'*x6');
    end;

 end;
 cmbbxObjetivos.Items.Clear;
 cmbbxobjetivos.Enabled := False;
 grdfuncaoobjetivo.Enabled := False;
 btnsalvaretapa3.Enabled := False;
 //btnredefiniretapa3.Enabled := True;
 btnexecutar.Enabled := True;
end
else
begin
 btnsalvaretapa3.Enabled := True;
 //btnredefiniretapa3.Enabled := False;
 btnexecutar.Enabled := False;
end;
end;
end;

procedure TIWForm2.btnBackRU2Click(Sender: TObject);
begin
iwrgnCaixaDialogoRU2.visible := false;
end;

procedure TIWForm2.btnBackRUClick(Sender: TObject);
begin
iwrgnCaixaDialogoRU.Visible := false;
end;

procedure TIWForm2.btnContinueregistClick(Sender: TObject);
var
i, j, auxcount : integer;
user_ ,datafinal: string;
begin
LimpaVariaveis;
With TIWUserSession(WebApplication.Data) do
begin
lnkplanos.Visible := False;
  with UserSession.ZQuery1 do
  begin
    // inicialmente contar quantos problemas o mesmo usuario já cadastrou
    SQL.Text := 'SELECT COUNT(name) FROM problem WHERE user = "' +
      (ID_user) + '";';
    ExecSQL;
    open;
    auxcount := FieldByName('COUNT(name)').AsInteger;
    close;

    { após saber quantos problemas o mesmo usuario cadastrou é possivel fazer um
    laço de repetição variando de zero ate essa quantidade-1 para pegar os nomes dos problemas }
    // VERIFICAR COMO FAZER ESSE PROCEDIMENTO

    i := 0;
    j := 0;
    SQL.Text := 'SELECT * FROM problem';
    ExecSQL;
    open;
    while i < auxcount do
    begin
      user_ := FieldByName('user').AsString;
      datafinal := FieldByName('conclusion_date').AsString;

      if ID_user = user_ then
      begin
        Setlength(nameproblem, j + 1);
        Setlength(idproblemasuser,j + 1);
        Setlength(datainicioproblema, j + 1);
        Setlength(datafinalproblema, j + 1);

        nameproblem[j] :=FieldByName('name').AsString;
        idproblemasuser[j] := FieldByName('id_problem').AsString;
        datainicioproblema[j] := FieldByName('creation_date').AsString;
        datafinalproblema[j] := FieldByName('conclusion_date').AsString;
        if datafinalproblema[j] = '' then
        datafinalproblema[j] := '02/01/1900';

        INC(j);
        // end;
        INC(i);
      end;
      next;
    end;
    close;
  end;

  if length(nameproblem) = 0 then
  begin
    WebApplication.ShowMessage('There is no problem to be continued');
    Exit;
  end;

  rgnopcaoinput.Visible := false;
  rgnretomada.Visible := true;
  btnauxchoose.Enabled := true;
  Retomada := True;
  problemapadrao := false;

  // colocar o vetor do nome no listbox
  lstbxProjects.Items.Clear;
  for i := 0 to length(nameproblem) - 1 do
  begin
    lstbxProjects.Items.Add(nameproblem[i] + ' - (' + datainicioproblema[i] + ')');
  end;
end;
end;


procedure TIWForm2.btncontinueRU2Click(Sender: TObject);
begin
WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm2.btncontinueRUClick(Sender: TObject);
begin
WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm2.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  quantregis: integer;

begin
  With TIWUserSession(WebApplication.Data) do
  begin
  lblwarningemail.Visible := false;
  lblwarningpassword.Visible := false;
  primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM user WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end;
  if quantregis = 0 then
  begin
    lblwarningemail.Visible := true;
    //IWButton1.Visible := true;
  end
  else
  begin
    email := edtemail.Text;
    with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      validation := FieldByName('validacao').AsString;
      close;

      SQL.Text := 'SELECT password FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      senha := FieldByName('password').AsString;
      close;
    end;
    if senha = edtSenha.Text then
    begin
      with UserSession.ZQuery1 do
      begin
        SQL.Text := 'SELECT id FROM user WHERE email="' + (edtemail.Text) + '"';
        ExecSQL;
        open;
        ID_user := FieldByName('id').AsString;
        close;

        SQL.Text := 'SELECT analyst FROM user WHERE email="' + (edtemail.Text)
          + '"';
        ExecSQL;
        open;
        analyst := FieldByName('analyst').AsInteger;
        close;

        if analyst = 0 then
        begin
          email := edtemail.Text;
          senha := edtSenha.Text;
        end;

      end;

      // * Caso o usuário esteja trocando a senha de cadastro *//
      {if validation = 'password' then
      begin
        WebApplication.ShowMessage('Temporary password has been found. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with changing your password.');
        edtSenha.Text := '';
        //Exit;
      end;}

      // * Caso seja o primeiro acesso *//
      if validation <> 'validado' then
      begin
        WebApplication.ShowMessage('User has not been validated. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with user validation.');
        edtSenha.Clear;
        Exit;
      end;

      rgnlogin.Visible := false;
      rgnopcaoinput.Visible := true;
    end
    else
    begin
      lblwarningpassword.Visible := true;
      //IWButton1.Visible := true;
    end;

  end;

 { // ** Salvando informação de data e horário
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 46;
  SalvarBD;

  // end;
  }
  end;
end;

procedure TIWForm2.btnexecutarClick(Sender: TObject);
var
    i, j, k, z1, z2, zj, z, t, counter1, counter2, aux2: integer;
    Matriz: Array of Array of Real;
    finished : boolean;
    violacao : boolean;
    violacao2 : array of boolean;
    duplicado : array of boolean;

    function Roleta: Real;
    begin
      // ** Roda obtendo um valor aleatório [0;100] para tomada de decisão **//
      Result := Random * 100;
    end;


    procedure PrimeiraPopulacao;
    var
      str1, str2: string;
      i, j, k, z1, z2, z: integer;
      NFuncAtual, EstoqueAtual, ProducaoTotalPeriodo: Real;
      AuxC, AuxD, AuxE, AuxEstoque, AuxV, AuxEE: Real;
      MaxProd: Array [0 .. 2] of Real;
      flag: boolean;
      valoranterior : array of double;
    begin
      // ** Gera a primeira população aleatoriamente **//
      { ** Cada indivíduo da população é uma matriz estruturada da seguinte forma:
        C1 : Número de ítens em produção regular
        C2: Número de ítens Prod em Horas Extra
        C3: Número de ítens de prod subcontratadas
        C4: Número de prod que ficaram em estoque no período anterior
        C5: Número de contratados no período atual
        C6: Número de demitidos no período atual
        C7: Número de funcionários remanecentes no período atual

        Cada indivíduo é avaliado com respeito às funções objetivo e possui uma linha
        na matriz PopulationPerformance que é estruturada da seguinte forma:

        C1: Performance no objetivo custo
        C2: Performance no objetivo credibilidade
        C3: Performance no objetivo Qualidade
        C4: Performance no objetivo Flexibilidade
        C5: Performance no objetivo rapidez
        C6: Demanda não atendida avaliada a cada período *** }

      // ** Define uma semente aleatória com base na hora **//

      With TIWUserSession(WebApplication.Data) do
      begin
      str1 := TimeToStr(Now);
      str2 := '';

      for i := 1 to Length(str1) - 1 do
        if str1[i] <> ':' then
          str2 := str2 + str1[i];

      RandSeed := 54730930; // strToInt(str2);

      // ** Define o tamanho da população como sendo 200 indivíduos **//
      Setlength(Populacao, 500, Periodo, 7);
      Setlength(Population, 500, Periodo, 7);
      setlength(valoranterior, 3);

      // ** Inicialmente os limites máximos de produção são :**//
      NFuncAtual := NinicialFunc;
      EstoqueAtual := 0;
      ProducaoTotalPeriodo := 0;
      if DemandaMedia <= Limites[0]* Limites[4] then
      MaxProd[0] := DemandaMedia // **Regular --> Atribui a média da demanda como sendo o máximo
      else
      MaxProd[0] := Limites[0]* Limites[4];
      MaxProd[1] := Limites[1] * Limites[4]; // **HoraExtra
      MaxProd[2] := Limites[2]; // **Subcontratação

      for j := 0 to Length(demanda) - 1 do
      begin
        if maxprod[0] + maxprod[1] + maxprod[2] < demanda[j] then
        begin
          maxprod[0] := Limites[0]* Limites[4];
          break;
        end;
      end;




      i := n1;
      while i <= Length(Populacao) - 1 do
      begin
       NFuncAtual := NinicialFunc;
       EstoqueAtual := 0;
       ProducaoTotalPeriodo := 0;
        // ** Períodos **//
        j := 0;
        while j <= (Length(Populacao[i]) - 1) do
        begin
          flag := false;
          if EstoqueAtual >= Demanda[j] then
            flag := true;

           if demanda[j] > (Limites[0]*limites[4] + Limites[1]*limites[4] + Limites[2] + EstoqueAtual)  then
           begin
             Populacao[i,j,0] := Limites[0]*limites[4];
             Populacao[i,j,1] := Limites[1]*limites[4];
             Populacao[i,j,2] := Limites[2];
             Populacao[i,j,3] := 0;
             Populacao[i,j,4] := limites[4] - NFuncAtual;
             Populacao[i,j,5] := 0;
             Populacao[i,j,6] := limites[4];
             inc(j);
             Continue;
           end;

          for k := 0 to 2 do
          valoranterior[k] := 0;

          // ** Evita gerar soluções inviáveis na primeira população **//
          while (ProducaoTotalPeriodo + EstoqueAtual <
          TIWUserSession(WebApplication.Data).Demanda[j]) or
            (TIWUserSession(WebApplication.Data).Populacao[i, j, 3] >
            TIWUserSession(WebApplication.Data).Limites[3]) or (flag) do
          begin


            ProducaoTotalPeriodo := 0;
            flag := false;
            for k := 0 to Length(Populacao[i, j]) - 1 do
            // ** Gerando plano aleatório para o período **//
            begin
              // ** Produção **//
              if k <= 2 then
              begin
                if k = 0 then
                begin
                  AuxE := MaxProd[k];
                end;
                if k = 2 then
                begin
                  AuxE := MaxProd[k];
                end;
                if k = 1 then
                begin
                if limites[0] <> 0 then
                  AuxV := (Populacao[i, j, 0] / Limites[0]);
                  if AuxV < NFuncAtual then
                    AuxV := NFuncAtual;

                  // ** a máxima produção em HE vai depender da quantidade de pessoal em prod regular **//
                  // ** Ninguém contrata pessoas apenas para trabalhar em hora extra **//
                  AuxE := AuxV * Limites[1];
                end;

                TIWUserSession(WebApplication.Data).Populacao[i, j, k] := Round(Random * AuxE);
                if populacao[i,j,k] < valoranterior[k] then
                TIWUserSession(WebApplication.Data).Populacao[i, j, k] := valoranterior[k];

                // ** Introduz uma "Mutação" das soluções em 25% dos casos **//
                if k = 0 then
                begin
                  // ** Força a produção da máxima capacidade regular **//
                  if Roleta < 25 then
                    Populacao[i, j, k] := NFuncAtual * Limites[0];

                  // ** Usar ao menos 60% da capacidade do sistema **//
                  if (TIWUserSession(WebApplication.Data).Populacao[i, j, k] < Limites[0] * NFuncAtual * 0.6) and (Demanda[j] >= Limites[0] * NFuncAtual * 0.6) then
                    Populacao[i, j, k] := Round(Limites[0] * NFuncAtual * 0.6);
                end;

                // ** Contabiliza a produção **//
                ProducaoTotalPeriodo := ProducaoTotalPeriodo + Populacao
                  [i, j, k];

                // ** Continua para a próxima iteração **//
                continue;
              end;

              // **Estoque **//
              if k = 3 then
              begin
                AuxC := (ProducaoTotalPeriodo + EstoqueAtual) -
                TIWUserSession(WebApplication.Data).Demanda[j];
                // ** A produção não atendeu a demada **//
                if AuxC < 0 then
                  begin
                  if Populacao[i, j, 0] <= limites[0]*Nfuncatual  then
                  valoranterior[0] := Populacao[i, j, 0];
                  if Populacao[i, j, 1] <= limites[1]*Nfuncatual  then
                  valoranterior[1] := Populacao[i, j, 1];
                  if Populacao[i, j, 2] <= limites[2] then
                  valoranterior[2] := Populacao[i, j, 2];

                  if valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual > demanda[j] then
                  begin
                    valoranterior[0] := valoranterior[0] - (valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual - demanda[j]);
                    if valoranterior[0] < 0 then
                    valoranterior[0] := 0;
                  end;

                  if valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual > demanda[j] then
                  begin
                    valoranterior[1] := valoranterior[1] - (valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual - demanda[j]);
                    if valoranterior[1] < 0 then
                    valoranterior[1] := 0;
                  end;

                  if valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual > demanda[j] then
                  begin
                    valoranterior[2] := valoranterior[2] - (valoranterior[0] + valoranterior[1] + valoranterior[2] + estoqueatual - demanda[j]);
                    if valoranterior[2] < 0 then
                    valoranterior[2] := 0;
                  end;


                  end;
                // ** Se já for o último período, não devemos ter estoque remanescente **//
                if j = Length(Populacao[i]) - 1 then
                  if AuxC <> 0 then
                  begin

                    // ** força uma "mutação" da solução **//
                    // ** Se o estoque já suprir a demanda nada será produzido **//
                    if EstoqueAtual >= Demanda[j] then
                    begin
                      Populacao[i, j, 0] := 0;
                      Populacao[i, j, 1] := 0;
                      Populacao[i, j, 2] := 0;
                      ProducaoTotalPeriodo := 0;
                      AuxC := EstoqueAtual - Demanda[j];

                    end
                    else if AuxC > 0 then
                    begin
                      ProducaoTotalPeriodo := 0;
                      // ** Caso contrário, deveremos produzir só o necessário para atender a demanda **//
                      // **Melhor dos casos **//
                      AuxD := Demanda[j] - EstoqueAtual;
                      AuxE := AuxD;
                      for z := 0 to 3 - 1 do
                      begin
                        // ** Se a Qtd a ser produzida já foi obtida, a produção nesta modalidade é nula **//
                        if AuxE = 0 then
                        begin
                          Populacao[i, j, z] := 0;
                          continue;
                        end;

                        // ** Verificando o valor a ser multiplicado pelo número aleatório **//
                        if MaxProd[z] > AuxE then
                          AuxEstoque := AuxE
                        else
                          AuxEstoque := MaxProd[z];

                        Populacao[i, j, z] := Round(Random * AuxEstoque);

                        // ** Desconta a quantidade já produzida **//
                        AuxE := AuxE - Populacao[i, j, z];

                        ProducaoTotalPeriodo := ProducaoTotalPeriodo + Populacao
                          [i, j, z];
                      end; // ZZ
                      // ** Caso ainda haja demanda **//
                      if AuxE > 0 then
                      begin
                        if Populacao[i, j, 0] + AuxE < NFuncAtual * Limites[0]
                          then
                          Populacao[i, j, 0] := Populacao[i, j, 0] + AuxE
                        else if Populacao[i, j, 1] + AuxE < NFuncAtual * Limites
                          [1] then
                          Populacao[i, j, 1] := Populacao[i, j, 1] + AuxE
                        else if Populacao[i, j, 2] + AuxE < NFuncAtual * Limites
                          [3] then
                          Populacao[i, j, 2] := Populacao[i, j, 2] + AuxE;
                        ProducaoTotalPeriodo := ProducaoTotalPeriodo + AuxE;
                      end;

                      AuxC := (ProducaoTotalPeriodo + EstoqueAtual)
                        - Demanda[j];
                    end; // Aux>0
                  end; // Aux <> 0
              end;

              // **Contratações**//
              if k = 4 then
              begin
                // ** Identifica a maior necessidade de funcionários **//
                if TIWUserSession(WebApplication.Data).Populacao[i, j, 0] / Limites[0] >= Populacao[i, j, 1]
                  / Limites[1] then
                  AuxD := Populacao[i, j, 0] / TIWUserSession(WebApplication.Data).Limites[0]
                else
                  AuxD := Populacao[i, j, 1] / Limites[1];

                // ** Armazena a informação para a próxima iteração **//
                AuxE := AuxD;
                AuxC := 0;
                // ** Contabiliza a necessidade de contratação **//
                while (AuxD > NFuncAtual) do
                begin
                  AuxD := AuxD - 1;
                  AuxC := AuxC + 1;
                end;
              end;

              // ** Demissões **//
              if k = 5 then
              begin
                // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
                AuxC := -1;
                // ** Contabiliza a necessidade de contratação **//
                while (AuxE <= NFuncAtual) do
                begin
                  AuxE := AuxE + 1;
                  AuxC := AuxC + 1;
                end;

                if AuxC < 0 then
                  AuxC := 0;

                // ** Se houver excesso de Mão de Obra **//
                if AuxC > 0 then
                begin
                  // ** Define aleatoriamente se haverá demissão **//
                  if (Roleta > 50) and (j <> Length(Populacao[i]) - 1) then
                    // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
                    AuxD := AuxC
                  else
                    AuxD := 0;

                  AuxC := AuxD
                end;
              end;

              // ** Total de funcionários remanescentes **//
              if k = 6 then
                AuxC := NFuncAtual + Populacao[i, j, 4] - Populacao[i, j, 5];

              Populacao[i, j, k] := AuxC;
            end; // ** k
          end; // **With



          //Garantindo o cumprimento de todas as restrições
          {if (Populacao[i, j, 3] > Limites[3]) or
          (Populacao[i, j, 0] > Limites[0]*Limites[4]) or
          (Populacao[i, j, 1] > Limites[1]*Limites[4]) or
          (Populacao[i, j, 2] > Limites[2]) or (NfuncAtual > Limites[4]) or
          ((j = Length(Populacao[i]) - 1) and (EstoqueAtual <> 0)) then
          begin
          if j <> 0 then
          begin
            NFuncAtual := Populacao[i, j-1, 6];
            EstoqueAtual := Populacao[i, j-1, 3];
            ProducaoTotalPeriodo := 0;
            dec(j);
          end
          else
          begin
            NFuncAtual := NinicialFunc;
            EstoqueAtual := 0;
            ProducaoTotalPeriodo := 0;
            dec(j);
          end;
          end; }

         // ** Caso a solução tenha sido adequada para o período, atualiza as var **//
          NFuncAtual := Populacao[i, j, 6];
          EstoqueAtual := Populacao[i, j, 3];
          ProducaoTotalPeriodo := 0;


          inc(j)
        end; // ** j
        violacao := false;
        for j := 0 to Length(Populacao[0]) - 1 do
        begin
          if (Populacao[i, j, 3] > Limites[3]) or (Populacao[i, j, 0] > Limites[0]*Limites[4]) or
            (Populacao[i, j, 1] > Limites[1]*Limites[4]) or (Populacao[i, j, 2] > Limites[2]) or (Populacao[i, j, 6] > Limites[4])
             or (Populacao[i, Length(Populacao[0]) - 1, 3] <> 0) then
           begin
             violacao := True;
             break;
           end;
        end;
        if violacao = false then
        inc(i)
      end; // i

      setlength(violacao2, Length(TIWUserSession(WebApplication.Data).Populacao));
      // ** Passando a populacao para o vetor que armazena o conjunto total de indivíduos a serem avaliados **//
      for i := 0 to Length(TIWUserSession(WebApplication.Data).Populacao) - 1 do
        for j := 0 to Length(Populacao[i]) - 1 do
          for k := 0 to Length(Populacao[i, j]) - 1 do
          begin
            Population[i, j, k] := Populacao[i, j, k];

            if (TIWUserSession(WebApplication.Data).Population[i, j, 3] > Limites[3]) or (Population[i, j, 0] > Limites[0]*Limites[4]) or
            (Population[i, j, 1] > Limites[1]*Limites[4]) or (Population[i, j, 2] > Limites[2]) or (Population[i, j, 6] > Limites[4])
             or (Population[i, Length(Population[0]) - 1, 3] <> 0) then
          violacao2[i] := True
          else
          violacao2[i] := false;
          end;
    end;
    end;

    procedure avaliarPopulacao;
    var
      i, j, k, z: integer;
      AuxCalc, AuxV, Prod: Real;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      Setlength(PopulationPerformance, Length(Population), length(objSelecionados)+1);

      for i := 0 to Length(Population) - 1 do
      begin
        PopulationPerformance[i, length(objSelecionados)] := 0;
        AuxV := 0;
        for j := 0 to length(objSelecionados)-1 do
        begin
          PopulationPerformance[i, j] := 0;
          AuxCalc := 0;

          for z := 0 to Length(Population[i]) - 1 do
          begin
            Prod := 0;
            // ** Custo **//
            if j = 0 then
            begin
              for k := 0 to Length(Population[i, z]) - 2 do
              begin
                PopulationPerformance[i, j] := PopulationPerformance[i, j]
                  + Population[i, z, k] * Custos[k];

                // ** Armazena a produção no período **//
                if k < 3 then
                  Prod := Prod + Population[i, z, k];

                // ** Avaliando atendimento a restrição de capacidade **//
                if k < 2 then
                  if Population[i, z, k] > Limites[k] * 10 then
                    AuxV := AuxV + Population[i, z, k] - Limites[k] * 10;
                if k = 2 then
                  if Population[i, z, k] > Limites[k] then
                    AuxV := AuxV + Population[i, z, k] - Limites[k];
              end;
              if z > 0 then
                // ** Adiciona o estoque remanescente no período anterior **//
                Prod := Prod + Population[i, z - 1, 3];
              AuxCalc := 1;
              // ** Avaliando o atendimento às restrições - Demanda **//
              if Demanda[z] > Prod then
                AuxV := AuxV + Demanda[z] - Prod;
              // ** Avaliando o atendo=imento a restrição de limite de estoque **//
              if Population[i, z, 3] > Limites[3] then
                AuxV := AuxV + Population[i, z, 3] - Limites[3];
            end
            else
            begin
              for k := 0 to 5 do
              begin
                PopulationPerformance[i, j] := PopulationPerformance[i, j]
                  + Population[i, z, k] * ParametrosFuncaoObjetivo[j - 1, k];
                if ParametrosFuncaoObjetivo[j - 1, k] <> 0 then
                AuxCalc := AuxCalc + Population[i, z, k];
              end;

            end; // ** else
          end; // ** z

          if AuxCalc = 0 then
            AuxCalc := 1;

          // ** Valor médio do objetivo **//
          PopulationPerformance[i, j] := PopulationPerformance[i, j] / AuxCalc;


        end; // ** J
        // ** Violação total da restrição **//
       // PopulationPerformance[i, length(objSelecionados)] := AuxV;
      end;

    end;
    end;

   procedure Crossover(X1, X2, z1 : integer);
    var
      i, j, k: integer;
      Aleatorio, Beta, Nc, AuxC, AuxD, AuxE1, AuxE2: Real;
      Producao, FunCAtual, Demissao: Array [0 .. 1] of Real;
      AtendeuDemanda : boolean;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      Nc := 20;

      AtendeuDemanda := false;
      while AtendeuDemanda = false do
      begin
      Aleatorio := Random;


      if Aleatorio <= 0.50 then
        Beta := 2 * Power((Aleatorio), (1 / (Nc + 1)))
      else
        Beta := (1 / (2 * Power((1 - Aleatorio), (1 / (Nc + 1)))));

      // ** Funcionários no início **//
      FunCAtual[0] := NinicialFunc;
     // FunCAtual[1] := NinicialFunc;
      //violacao := false;
      i := 0;
      while i <= Length(NovaGeracao[0]) - 1 do
      begin
        Producao[0] := 0;
       // Producao[1] := 0;

        // ** Valores da produ~ção **//
        for j := 0 to 3 - 1 do
        begin

          // ** Filho 1 **//
          NovaGeracao[z1, i, j] := ABS
            (Round(0.5 * ((1 + Beta) * Populacao[X1, i, j] + (1 - Beta)
                  * Populacao[X2, i, j])));
          // ** Filho 22 **//
          {NovaGeracao[z2, i, j] := ABS
            (Round(0.5 * ((1 - Beta) * Populacao[X1, i, j] + (1 + Beta)
                  * Populacao[X2, i, j]))); }

          // ** Coletando dados **//
          Producao[0] := Producao[0] + NovaGeracao[z1, i, j];
        //  Producao[1] := Producao[1] + NovaGeracao[z2, i, j];
        end;

        // ** Calculando o estoque do período **//
        if i > 0 then
          AuxC := NovaGeracao[z1, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[0] < 0) then
          begin
          NovaGeracao[z1, i, 3] := AuxC + Producao[0] - Demanda[i];
          AtendeuDemanda := true;
          end;
        if (Demanda[i] - AuxC - Producao[0] = 0) then
          NovaGeracao[z1, i, 3] := 0;


        if (Demanda[i] - AuxC - Producao[0] > 0) then   //Não atendeu a demanda
        begin
         AtendeuDemanda := false;
         i := Length(NovaGeracao[0]);
         continue;
        end;


        {if i > 0 then
          AuxC := NovaGeracao[z2, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[1] < 0) then
          NovaGeracao[z2, i, 3] := AuxC + Producao[1] - Demanda[i]
        else
          NovaGeracao[z2, i, 3] := 0;        }
        /// ////////////////////////////////////////////////////

        // ** Calculando contratações **//
        if NovaGeracao[z1, i, 0] / Limites[0] >= NovaGeracao[z1, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z1, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z1, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE1 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[0]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z1, i, 4] := AuxC;
        FunCAtual[0] := FunCAtual[0] + AuxC;

        {// ** Calculando contratações **//
        if NovaGeracao[z2, i, 0] / Limites[0] >= NovaGeracao[z2, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z2, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z2, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE2 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[1]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z2, i, 4] := AuxC;
        FunCAtual[1] := FunCAtual[1] + AuxC;  }
        /// ////////////////////////////////////////////////////////////////////

        // ** Demissões **//
        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE1 <= FunCAtual[0]) do
        begin
          AuxE1 := AuxE1 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z1, i, 5] := AuxC;
        FunCAtual[0] := FunCAtual[0] - AuxC;

       { // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE2 <= FunCAtual[0]) do
        begin
          AuxE2 := AuxE2 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z2, i, 5] := AuxC;
        FunCAtual[1] := FunCAtual[1] - AuxC;  }

        // ** Total de funcionários remanescentes **//
        NovaGeracao[z1, i, 6] := FunCAtual[0];
        //NovaGeracao[z2, i, 6] := FunCAtual[1];



      inc(i);
      end; // ** i

      end; // Atendeu a Demanda

    end;
    end;

    procedure Mutation(X1, X2, z1: integer);
    var
      i, j, k: integer;
      Aleatorio, Beta, Nc, AuxC, AuxD, AuxE1, AuxE2: Real;
      Producao, FunCAtual, Demissao: Array [0 .. 1] of Real;
      AtendeuDemanda : boolean;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      // ** Funcionários no início **//
      FunCAtual[0] := NinicialFunc;
      //FunCAtual[1] := NinicialFunc;

      AtendeuDemanda := false;
      while AtendeuDemanda = false do
      begin
         i := 0;

      while i <= Length(NovaGeracao[0]) - 1 do
      begin
        Producao[0] := 0;
       // Producao[1] := 0;



        // ** Valores da produ~ção **//
        for j := 0 to 3 - 1 do
        begin
          // ** Mutação 1 **//
          if Roleta > 50 then
            Aleatorio := Random
          else
            Aleatorio := 1;

          NovaGeracao[z1, i, j] := Round(TIWUserSession(WebApplication.Data).Populacao[X1, i, j] * 0.5 + Populacao[X1, i, j] * Aleatorio);

        {  // ** Mutação 2 **//
          NovaGeracao[z2, i, j] := Round
            (Populacao[X2, i, j] * 0.5 + Aleatorio * Populacao[X2, i, j]);  }

          // ** Coletando dados **//
          Producao[0] := Producao[0] + NovaGeracao[z1, i, j];
         // Producao[1] := Producao[1] + NovaGeracao[z2, i, j];
        end;

        // ** Calculando o estoque do período **//
        if i > 0 then
          AuxC := NovaGeracao[z1, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[0] < 0) then
        begin
          NovaGeracao[z1, i, 3] := AuxC + Producao[0] - Demanda[i];
          AtendeuDemanda := true;
        end;
        if (Demanda[i] - AuxC - Producao[0] = 0) then
          NovaGeracao[z1, i, 3] := 0;
        if (Demanda[i] - AuxC - Producao[0] > 0) then  // Não atendeu a demanda
        begin
          AtendeuDemanda := false;
          i := Length(NovaGeracao[0]);
          Continue;
        end;

       { if i > 0 then
          AuxC := NovaGeracao[z2, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[1] < 0) then
          NovaGeracao[z2, i, 3] := AuxC + Producao[1] - Demanda[i]
        else
          NovaGeracao[z2, i, 3] := 0;
        /// ////////////////////////////////////////////////////
        }
        // ** Calculando contratações **//
        if NovaGeracao[z1, i, 0] / Limites[0] >= NovaGeracao[z1, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z1, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z1, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE1 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[0]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z1, i, 4] := AuxC;
        FunCAtual[0] := FunCAtual[0] + AuxC;

      {  // ** Calculando contratações **//
        if NovaGeracao[z2, i, 0] / Limites[0] >= NovaGeracao[z2, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z2, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z2, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE2 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[1]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z2, i, 4] := AuxC;
        FunCAtual[1] := FunCAtual[1] + AuxC;
       } /// ////////////////////////////////////////////////////////////////////

        // ** Demissões **//
        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE1 <= FunCAtual[0]) do
        begin
          AuxE1 := AuxE1 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z1, i, 5] := AuxC;
        FunCAtual[0] := FunCAtual[0] - AuxC;

       { // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE2 <= FunCAtual[0]) do
        begin
          AuxE2 := AuxE2 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z2, i, 5] := AuxC;
        FunCAtual[1] := FunCAtual[1] - AuxC;

        // ** Total de funcionários remanescentes **//
        NovaGeracao[z1, i, 6] := FunCAtual[0];
        NovaGeracao[z2, i, 6] := FunCAtual[1];
      }

     { j := i;
      if (NovaGeracao[z1, j, 3] > Limites[3]) or (NovaGeracao[z1, j, 0] > Limites[0]*Limites[4]) or
            (NovaGeracao[z1, j, 1] > Limites[1]*Limites[4]) or (NovaGeracao[z1, j, 2] > Limites[2]) or (NovaGeracao[z1, j, 6] > Limites[4])
             or (NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] <> 0) then
           begin
             violacao[z1] := True;
             //break;
           end;
        if (NovaGeracao[z2, j, 3] > Limites[3]) or (NovaGeracao[z2, j, 0] > Limites[0]*Limites[4]) or
            (NovaGeracao[z2, j, 1] > Limites[1]*Limites[4]) or (NovaGeracao[z2, j, 2] > Limites[2]) or (NovaGeracao[z2, j, 6] > Limites[4])
             or (NovaGeracao[z2, Length(NovaGeracao[0]) - 1, 3] <> 0) then
           begin
             violacao[z2] := True;
             //break;
           end;
                }

      inc(i);

      end; // ** i

      end;

    end;
    end;
    function RankearPopulacao: boolean;
    var
      i, j, k, z, n, aux2, ii, jj: integer;
      Ndominancias: Array of integer; // ** Número de projetos que domina cada projeto
      NSup: Array [0 .. 1] of integer;
      aux1, aux3 : string;
      pula : boolean;
      Resultado : array of real;
    begin
     With TIWUserSession(WebApplication.Data) do
      begin
      SetLength(PerformanceNormalizada, Length(PopulationPerformance), length(PopulationPerformance[0])-1);
      SetLength(Maior, length(PopulationPerformance[0])-1);
      SetLength(Menor, length(PopulationPerformance[0])-1);


      for j := 0 to length(PopulationPerformance[0])-2 do
      begin
        Maior[j] := -99999999999999;
        Menor[j] := 999999999999999;
        for i := 0 to Length(PopulationPerformance)-1 do
        begin
        if maior[j] < PopulationPerformance[i,j] then
        maior[j] := PopulationPerformance[i,j];
        if menor[j] > PopulationPerformance[i,j] then
        menor[j] := PopulationPerformance[i,j];
        end;
      end;


      for i := 1 to Length(PopulationPerformance) do
      for j := 1 to length(PopulationPerformance[i])-1 do
      begin
        if objSelecionados[j-1] <> 'Cost' then
          TIWUserSession(WebApplication.Data).PerformanceNormalizada[i - 1, j - 1] :=(PopulationPerformance[i - 1, j - 1] - TIWUserSession(WebApplication.Data).menor[j - 1]) /
            (maior[j - 1] - menor[j - 1])
        else if objSelecionados[j-1] = 'Cost' then
          PerformanceNormalizada[i - 1, j - 1] :=(TIWUserSession(WebApplication.Data).PopulationPerformance[i - 1, j - 1] - maior[j - 1]) /
          (menor[j - 1] - maior[j - 1]);
      end;
      SetLength(Resultado, Length(PerformanceNormalizada));
      for i := 0 to Length(PerformanceNormalizada) - 1 do
      begin
        Resultado[i] := 0;
        for J := 0 to length(PerformanceNormalizada[i])-1 do
        Resultado[i] := Resultado[i] + PerformanceNormalizada[i,j];
      end;

      SetLength(duplicado, Length(TIWUserSession(WebApplication.Data).Population));
      for i := 0 to Length(Population) - 1 do
      duplicado[i] := False;

      for i := 0 to Length(Population) - 1 do
      Begin
        for j := i + 1 to Length(Population) - 1 do
        begin
         aux2 := 0;
         ii := 0;
         jj := 0;
         pula := false;
         for ii := 0 to Length(PopulationPerformance[i]) - 1 do
         begin
           if pula = true then
           continue;
           if PopulationPerformance[i,ii] = PopulationPerformance[j,ii] then
            inc(aux2)
           else
           pula := true;
         end;

         if aux2 = Length(PopulationPerformance[0]) then
         duplicado[j] := true;

          if duplicado[j] = true then
         begin
           for k := 0 to length(PopulationPerformance[0]) - 1 do
           begin
             if (k = 0) or (k = length(PopulationPerformance[0]) - 1) then
             PopulationPerformance[j,k] := 99999999999999
             else
             PopulationPerformance[j,k] := -99999999999999;

           end;
         end;

        end;
      End;
      for i := 0 to Length(Population) - 1 do
      Begin
        if (duplicado[i] = true) or (violacao2[i] = true) then
        begin
        Resultado[i] := -999999999999;
        violacao2[i] := true;
        end;
      End;

       Setlength(PopulationRanking, 0, 0);
      Setlength(PopulationRanking, 1);
      Setlength(Ndominancias, Length(Population));

      // ** Inicializa a variável **//
      for i := 0 to Length(Population) - 1 do
      begin
        Ndominancias[i] := 0;
      end;
      // ** Comparação Par a Par **//

     for i := 0 to Length(Population) - 1 do
      Begin
        for j := i + 1 to Length(Population) - 1 do
        begin
        // ** Regras de dominância restrita **//
          if ((TIWUserSession(WebApplication.Data).PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              < PopulationPerformance[j, Length(PopulationPerformance[j]) - 1])) or
              ((violacao2[i] = false) and (violacao2[j] = true))
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[j]);
            continue;
          end;
          if ((PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              > PopulationPerformance[j, Length(PopulationPerformance[i]) - 1])) or ((violacao2[j] = false) and (violacao2[i] = true))
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[i]);
            continue;
          end;

          if Resultado[i] > Resultado[j]
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[j]);
            continue;
          end;
          if Resultado[j] > Resultado[i]
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[i]);
            continue;
          end;


          // ** Caso nenhuma das regras tenham sido atendidas, os projetos são incomparáveis **//
        end;
        // ** Caso I não seja dominado por ninguém, ele pertence ao R1 **//
        if Ndominancias[i] = 0 then
        begin
          Setlength(PopulationRanking[0], Length(PopulationRanking[0]) + 1);
          // ** inclui na primeira posição **//
          PopulationRanking[0, Length(PopulationRanking[0]) - 1] := i;
        end;
      End;
      z := 1;

      n := 0;
      while n < Length(Population) do
      begin
        n := 0;
        for j := 0 to Length(Population) - 1 do
        begin
          // ** Subtrai 1 no número de Dominâncias para a construção do ranking**//
          Dec(Ndominancias[j]);
          if Ndominancias[j] <= 0 then
            Inc(n);
        end;
        z := 0;

        for i := 0 to Length(Population) - 1 do
        begin
          // ** Se após reduzir em um, o num de dom se tornar zero, a sol pertence ao atual nível do Ranking **//
          if Ndominancias[i] = 0 then
          begin
            Inc(z);

            // ** Cria o nível atual do ranking **//
            if z = 1 then
              Setlength(PopulationRanking, Length(PopulationRanking) + 1);

            // ** Adiciona mais um slot ao nível atual **//
            Setlength(PopulationRanking[Length(PopulationRanking) - 1], Length
                (PopulationRanking[Length(PopulationRanking) - 1]) + 1);
            // ** Adiciona a solução ao atual nível **//
            PopulationRanking[Length(PopulationRanking) - 1, Length
              (PopulationRanking[Length(PopulationRanking) - 1]) - 1] := i;
          end;
        end; // **i

      end; // ** while z

      // ** Se todas os indivíduos forem não dominados **//
      if Length(PopulationRanking) = 1 then
        // ** Avisa que chegamos ao final do processo **//
        Result := true
      else
        Result := false;

    end;
    end;






    function RankingSorting: boolean;
    var
      i, j, k, z, n, aux2: integer;
      Ndominancias: Array of integer; // ** Número de projetos que domina cada projeto
      NSup: Array [0 .. 1] of integer;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      SetLength(duplicado, Length(TIWUserSession(WebApplication.Data).Population));
      for i := 0 to Length(Population) - 1 do
      duplicado[i] := False;

      for i := 0 to Length(Population) - 1 do
      Begin
        for j := i + 1 to Length(Population) - 1 do
        begin
         aux2 := 0;
         k := 0;
         while k <= Length(PopulationPerformance[i]) - 1 do
         begin
           if PopulationPerformance[i,k] = PopulationPerformance[j,k] then
            inc(aux2)
           else
             k := Length(PopulationPerformance[i]);
            inc(k);
         end;

         if aux2 = Length(PopulationPerformance[0]) then
         duplicado[j] := true;





        end;



        end;






      Setlength(PopulationRanking, 0, 0);
      Setlength(PopulationRanking, 1);
      Setlength(Ndominancias, Length(Population));

      // ** Inicializa a variável **//
      for i := 0 to Length(Population) - 1 do
      begin
        Ndominancias[i] := 0;
      end;
      // ** Comparação Par a Par **//




      for i := 0 to Length(Population) - 1 do
      Begin
        for j := i + 1 to Length(Population) - 1 do
        begin
         if duplicado[i] or violacao2[i] then
         Inc(Ndominancias[i]);

         if duplicado[j] or violacao2[j] then
         Inc(Ndominancias[j]);

         if duplicado[j] or violacao2[j] or duplicado[i] or violacao2[i] then
         continue;




          if ((TIWUserSession(WebApplication.Data).PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              < PopulationPerformance[j, Length(PopulationPerformance[j]) - 1]))
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[j]);
            continue;
          end;
          if ((PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              > PopulationPerformance[j, Length(PopulationPerformance[i]) - 1]))
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[i]);
            continue;
          end;
          // ** Inicializa contadores **//
          NSup[0] := 0;
          NSup[1] := 0;

          // ** Caso uma dominancia não tenha sido estabelecida ainda, vai para os payoffs **//
          for k := 0 to Length(PopulationPerformance[i]) - 2 do
            // ** Custo - Minimização **
            if k = 0 then
            begin
              if PopulationPerformance[i, k] > PopulationPerformance[j, k] then
                Inc(NSup[1]);
              if PopulationPerformance[i, k] < PopulationPerformance[j, k] then
                Inc(NSup[0]);
            end
            else
            begin
              // ** Obj de maximização **//
              if PopulationPerformance[i, k] > PopulationPerformance[j, k] then
                Inc(NSup[0]);
              if PopulationPerformance[i, k] < PopulationPerformance[j, k] then
                Inc(NSup[1]);
            end;

          // ** I Domina J **//
          if (NSup[0] > 0) and (NSup[1] = 0) then
            Inc(Ndominancias[j]);

          // ** J Domina I **//
          if (NSup[1] > 0) and (NSup[0] = 0) then
            Inc(Ndominancias[i]);

          // ** Caso nenhuma das regras tenham sido atendidas, os projetos são incomparáveis **//
        end;
        // ** Caso I não seja dominado por ninguém, ele pertence ao R1 **//
        if Ndominancias[i] = 0 then
        begin
          Setlength(PopulationRanking[0], Length(PopulationRanking[0]) + 1);
          // ** inclui na primeira posição **//
          PopulationRanking[0, Length(PopulationRanking[0]) - 1] := i;
        end;
      End;
      z := 1;

      n := 0;
      while n < Length(Population) do
      begin
        n := 0;
        for j := 0 to Length(Population) - 1 do
        begin
          // ** Subtrai 1 no número de Dominâncias para a construção do ranking**//
          Dec(Ndominancias[j]);
          if Ndominancias[j] <= 0 then
            Inc(n);
        end;
        z := 0;

        for i := 0 to Length(Population) - 1 do
        begin
          // ** Se após reduzir em um, o num de dom se tornar zero, a sol pertence ao atual nível do Ranking **//
          if Ndominancias[i] = 0 then
          begin
            Inc(z);

            // ** Cria o nível atual do ranking **//
            if z = 1 then
              Setlength(PopulationRanking, Length(PopulationRanking) + 1);

            // ** Adiciona mais um slot ao nível atual **//
            Setlength(PopulationRanking[Length(PopulationRanking) - 1], Length
                (PopulationRanking[Length(PopulationRanking) - 1]) + 1);
            // ** Adiciona a solução ao atual nível **//
            PopulationRanking[Length(PopulationRanking) - 1, Length
              (PopulationRanking[Length(PopulationRanking) - 1]) - 1] := i;
          end;
        end; // **i

      end; // ** while z

      // ** Se todas os indivíduos forem não dominados **//
      if Length(PopulationRanking) = 1 then
        // ** Avisa que chegamos ao final do processo **//
        Result := true
      else
        Result := false;

    end;
    end;

    procedure CrowdingDistance;
    var
      i, j, Katual, k, n: integer;
      OrderM: Array of integer;
      ConsA, ConsB, Mg, ConsRange: Real;
      Desempenho: array of Real;
      Ndominancias: Array of integer; // ** Número de projetos que domina cada projeto
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      Setlength(PopulationDistance, Length(Population));
      Setlength(Desempenho, Length(Population));
      Setlength(OrderM, Length(Population));

      // ** Inicializando **//
      for i := 0 to Length(Population) - 1 do
        PopulationDistance[i] := 0;

      // ** Para cada objetivo **//
      for j := 0 to Length(PopulationPerformance[0]) - 2 do
      begin

        // ** Passa o desempenho para o vet auxiliar **//
        for i := 0 to Length(PopulationPerformance) - 1 do
          Desempenho[i] := PopulationPerformance[i, j];

        Setlength(Ndominancias, Length(Population));

        // ** Inicializa a variável **//
        for i := 0 to Length(Population) - 1 do
        begin
          Ndominancias[i] := 0;
          OrderM[i] := -1;
        end; // **i

        // ** Ordena os indivíduos do menor p/ o maior**//
        for i := 0 to Length(Population) - 1 do
        begin
          for k := i to Length(Population) - 1 do
          begin
            if Desempenho[i] < Desempenho[k] then
              Inc(Ndominancias[k]);
            if Desempenho[i] > Desempenho[k] then
              Inc(Ndominancias[i]);
          end; // **k

        end; // **i

        Katual := 0;
        n := 0;
        ConsA := -1;
        ConsB := -1;

        // ** Ordena com base no número de dominâncias recebidas **//
        while OrderM[Length(Population) - 1] = -1 do
        begin
          for k := 0 to Length(Population) - 1 do
            // ** Já não existem indivíduos com desempenho menor que K;
            if Ndominancias[k] = n then
            begin
              OrderM[Katual] := k;
              Inc(Katual);
              // ** Seleciona o menor e a melhor performance **//
              if n = 0 then
                ConsA := Desempenho[k];
              if Katual = Length(Population) then
                ConsB := Desempenho[k];
            end; // **Ndom

          if ConsB = -1 then
            if OrderM[Katual] = -1 then
              Inc(n);
        end; // **While

        if ConsA = ConsB then
        begin
          Mg := 0;
          ConsRange := 1
        end
        else
        begin
          Mg := Power(10, 5);
          ConsRange := ConsB - ConsA;
        end;

        // ** Define o Crowding Distance **//
        for i := 1 to Length(Population) - 2 do
        begin
          // ** O primeiro e o último indivíduo da população **//
          if (i = 1) then
            PopulationDistance[OrderM[i - 1]] := PopulationDistance
              [OrderM[i - 1]] + Mg;
          if (i = Length(Population) - 2) then
            PopulationDistance[OrderM[i + 1]] := PopulationDistance
              [OrderM[i + 1]] + Mg;

          PopulationDistance[OrderM[i]] := PopulationDistance[OrderM[i]] +
            (Desempenho[OrderM[i + 1]] - Desempenho[OrderM[i - 1]]) / ConsRange;

          if PopulationDistance[OrderM[i]] < 0 then
            PopulationDistance[OrderM[i]] := ABS(PopulationDistance[OrderM[i]]);

        end;

      end; // ** j

    end;
    end;

    procedure SelecaoNatural;
    var
      i, j, Isel, k, z, ICandidato: integer;
      UltimoNivel: boolean;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      Isel := 0;

      Setlength(PerformancePopulacao, Length(Populacao), Length
          (PopulationPerformance[0]));
      Setlength(DistanciaPopulacional, Length(Populacao));
      Setlength(OrdemPopulacao, 0, 0);

      // ** Percorre os níveis do ranking **//
      for i := 0 to Length(PopulationRanking) - 1 do
      begin
        // ** Acrescenta um nível no ranking final **//
        Setlength(OrdemPopulacao, i + 1);
        if Isel + Length(PopulationRanking[i]) > Length(Populacao) then
          UltimoNivel := true
        else
          UltimoNivel := false;

        for j := 0 to Length(PopulationRanking[i]) - 1 do
        begin
          ICandidato := PopulationRanking[i, j];
          // ** Todo o nível não cabe na população **//
          if UltimoNivel then
          begin
            for k := 0 to Length(PopulationRanking[i]) - 1 do
              if PopulationDistance[ICandidato] < PopulationDistance
                [PopulationRanking[i, k]] then
                // ** Seleciona a maior distância Populacional **//
                ICandidato := PopulationRanking[i, k];
          end;

          // ** Armazena a solução **//
          for k := 0 to Length(Population[ICandidato]) - 1 do
            for z := 0 to Length(Population[ICandidato, k]) - 1 do
              Populacao[Isel, k, z] := Population[ICandidato, k, z];

          // ** Armazena a performance **//
          for k := 0 to Length(PerformancePopulacao[0]) - 1 do
            PerformancePopulacao[Isel, k] := PopulationPerformance
              [ICandidato, k];

          // ** Armazena a Crowding Distance **//
          DistanciaPopulacional[Isel] := PopulationDistance[ICandidato];

          // ** Acrescenta um slot no nível atual **//
          Setlength(OrdemPopulacao[i], Length(OrdemPopulacao[i]) + 1);

          // ** Amazena o Ranking remanescente **//
          OrdemPopulacao[i, Length(OrdemPopulacao[i]) - 1] := Isel;

          // ** A partir desse momento, altera o Crowding D para evitar que se repita alguma solução**//
          if UltimoNivel then
            PopulationDistance[ICandidato] := -100;

          Inc(Isel);

          // ** Finaliza o processo **//
          if Isel = Length(Populacao) then
            exit;
        end; // **j

      end; // **i

    end;
    end;

    Procedure SelecaoGenitores;
    var
      Rposition, i, j: integer;
      X: array [0 .. 1] of integer;
      RDAtual: Real;
    begin
      With TIWUserSession(WebApplication.Data) do
      begin
      G1 := 0;
      G2 := 0;

      // ** Os genitores devem ser distintos um do outro **//
      while G1 = G2 do
      begin
        for i := 0 to 2 - 1 do
        begin
          // ** Seleciona a posição do ranking aleatoriamente **//
          Rposition := Round((Length(OrdemPopulacao) - 1) * Random);

          // ** Como é bastante difícil [dependendo do num de niveis] selecionar o primeiro e o último níveis, damos uma mutada **//
          if Roleta < 30 then
            // ** +30% de chance de selecionar a primeira posição **//
            Rposition := 0
          else if Roleta > 95 then
            // ** +5% de chance de selecionar a ultima posição **//
            Rposition := Length(OrdemPopulacao) - 1;

          RDAtual := -8000;

          // **Seleciona aleatoriamente em 10% das vezes **//
          if Roleta < 10 then
            X[i] := Round((Length(OrdemPopulacao[Rposition]) - 1) * Random)
          else
            // ** Busca o indivíduo com maior crowding distance no ranking **//
            for j := 0 to Length(OrdemPopulacao[Rposition]) - 1 do
              if RDAtual < DistanciaPopulacional[OrdemPopulacao[Rposition, j]]
                then
              begin
                RDAtual := DistanciaPopulacional[OrdemPopulacao[Rposition, j]];
                X[i] := OrdemPopulacao[Rposition, j];
              end;
        end;
        /// ** i

        G1 := X[0];
        G2 := X[1];
      end; // **While

    end;
    end;

  begin
   btnexecutar.Enabled := False;
   btnreiniciaretapa3.Enabled := False;
  // rgnfuncaoobjetivo.Cursor := crWait;

    // ** Gera a primeira geração **//
    With TIWUserSession(WebApplication.Data) do
    begin
    Finalize(populacao);
    finalize(population);
    finalize(NovaGeracao);
    finalize(PerformancePopulacao);

    PrimeiraPopulacao;

    finished := false;

     Nger := 0;
    // ** Realiza o processo evolutivo **//
    while not finished do
    begin
      // ** Avalia o desempenho da população **//
      avaliarPopulacao;
      // ** Ranqueia as soluções **//
      finished := RankingSorting;
      // **Calcula o Crowding distance **//
      CrowdingDistance;
      // ** Realiza a seleção dos indivíduos que sobreviverão **//
      SelecaoNatural;

      // **Gerar uma nova populacao **//
      Setlength(TIWUserSession(WebApplication.Data).NovaGeracao, Length(
      TIWUserSession(WebApplication.Data).Populacao), Length(Populacao[0]), Length
          (Populacao[0, 0]));

      z1 := 0;
      //z2 := 1;

      // ** Criar uma nova população de tamanho padrão **//

      While z1 < Length(Populacao) do
      begin
        // ** Seleciona genitores **//
        SelecaoGenitores;


        // ** 70% de chance de crossover, 30% de mutação **//
        if Roleta < 70 then
        begin
          Crossover(G1, G2, z1);


        end
        else
        begin
          Mutation(G1, G2, z1);

        end;

        if (NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] <> 0) then
        begin
         if NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 0] >= NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] then
         begin
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 0] := NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 0] - NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3];
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] := 0;
         end;
         if NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 1] >= NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] then
         begin
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 1] := NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 1] - NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3];
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] := 0;
         end;
         if NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 2] >= NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] then
         begin
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 2] := NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 2] - NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3];
         NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] := 0;
         end;
        end;

        violacao := false;
        for j := 0 to Length(NovaGeracao[0]) - 1 do
        begin
          if (TIWUserSession(WebApplication.Data).NovaGeracao[z1, j, 3] > Limites[3]) then
          begin
             violacao := True;
             break;
           end;
          if (NovaGeracao[z1, j, 0] > Limites[0]*Limites[4]) then
          begin
             violacao := True;
             break;
           end;
          if  (NovaGeracao[z1, j, 1] > Limites[1]*Limites[4]) then
          begin
             violacao := True;
             break;
           end;
          if (NovaGeracao[z1, j, 2] > Limites[2]) then
          begin
             violacao := True;
             break;
           end;
          if (NovaGeracao[z1, j, 6] > Limites[4]) then
          begin
             violacao := True;
             break;
           end;
          if  (NovaGeracao[z1, Length(NovaGeracao[0]) - 1, 3] <> 0) then
          begin
             violacao := True;
             break;
           end;
          if ((NovaGeracao[z1,0,6]-NovaGeracao[z1,0,4] + NovaGeracao[z1,0,5]) <> TIWUserSession(WebApplication.Data).NinicialFunc) then
           begin
             violacao := True;
             break;
           end;
        end;
        if violacao = false then
        inc(z1);

      end; // ** z

      // **Junta a nova população com a anterior **//
      Setlength(Population, 2 * Length(TIWUserSession(WebApplication.Data).Populacao), Length(Populacao[0]), Length
          (Populacao[0, 0]));

      SetLength(violacao2, 2 * Length(TIWUserSession(WebApplication.Data).Populacao));

      for i := 0 to 2 * Length(Populacao) - 1 do
      begin
      violacao2[i] := false;
        for j := 0 to Length(Populacao[0]) - 1 do
          for k := 0 to Length(Populacao[0, 0]) - 1 do
          begin
            if i < Length(Populacao) then
              Population[i, j, k] := Populacao[i, j, k]
            else
            Population[i, j, k] := NovaGeracao[i - Length(TIWUserSession(WebApplication.Data).Populacao), j, k];


         if violacao2[i] = true then
         continue;

         if (TIWUserSession(WebApplication.Data).Population[i, j, 3] > TIWUserSession(WebApplication.Data).Limites[3])
         or (Population[i, j, 0] > TIWUserSession(WebApplication.Data).Limites[0]*Limites[4])
         or (Population[i, j, 1] > Limites[1]*Limites[4])
         or (Population[i, j, 2] > Limites[2])
         or (Population[i, j, 6] > Limites[4])
         or (Population[i, Length(Population[0]) - 1, 3] <> 0) then
          violacao2[i] := True;
          end;
      end;



      Inc(Nger);
      if TIWUserSession(WebApplication.Data).Nger = 1000 then
        finished := true;

    end; // ** Finished

  IwTimer1.Enabled := True;
  //Exportar;


  lnkplanos.Visible := True;
 SalvarBD;

  //Update;
  btnexecutar.Enabled := True;
   btnreiniciaretapa3.Enabled := True;
   TfrmResultados.Create(WebApplication).show;
   //rgnfuncaoobjetivo.Cursor := crDefault;
    end;
  end;

procedure TIWForm2.btnexportarClick(Sender: TObject);
var
i, counter2, k, z : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
    grdplanos.columns.Clear;
    counter2 := -1;
    grdplanos.TotalRows := 3;
    for I := 0 to 8 do
    begin
    grdplanos.columns.Add;
    end;
    for I := 0 to Length(Populacao) - 1 do
    begin
    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
    grdplanos.Cells[0,Counter2+1] := 'Plan ' + inttostr(i+1) + ':';
    //sheet2.Cells[Counter2+2, 1] := 'Plan ' + inttostr(i+1) + ':';
    for k := 0 to Length(Populacao[i]) - 1 do
    begin
    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
    //grdplanos.Cells[0,Counter2+1] := inttostr(z+1);
      for z := 0 to Length(Populacao[i, k]) - 1 do
        begin
       grdplanos.Cells[Z + 1,Counter2+1] := floatToStr(Populacao[i, k, z]);
       //sheet2.Cells[Counter2+2,z+2] := floatToStr(Populacao[i, k, z]);
        end;
    end;
    end;
    grdplanos.Cells[0,0] := 'Plan';
   grdplanos.Cells[1,0] := 'Regular Production';
   grdplanos.Cells[2,0] := 'Overtime Production';
   grdplanos.Cells[3,0] := 'Subcontrating Production';
   grdplanos.Cells[4,0] := 'Stock';
   grdplanos.Cells[5,0] := 'Hiring';
   grdplanos.Cells[6,0] := 'Resignation';
   grdplanos.Cells[7,0] := 'Total Employees';
TIWAdvWebGridExcelIO2.AdvWebGrid := grdplanos;
  TIWAdvWebGridExcelIO2.XLSExport('Plans_' + problema + '.xls', WebApplication);
  //Inherited;
end;
end;

procedure TIWForm2.Exportar;
var
  ExcelApplication: variant;
  Sheet1, Sheet2, Sheet3: variant;
  i, j, z, k, t, counter1, counter2: integer;
  Matriz: Array of Array of Real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  {try
    begin
      ExcelApplication := CreateOleObject('Excel.Application');
    end;
  except
    ShowMessage(
      'It was not possible to create an Excel file. Make sure that MS Excel is installed on your system.');
  end;
  ExcelApplication.WorkBooks.Add(1);

  ExcelApplication.WorkBooks[1].Sheets.Add;
  ExcelApplication.WorkBooks[1].Sheets.Add;
  ExcelApplication.WorkBooks[1].Sheets.Add;

  ExcelApplication.WorkBooks[1].Worksheets[1].Name := 'Alternatives';
  ExcelApplication.WorkBooks[1].Worksheets[2].Name := 'Plans';

  Sheet1 := ExcelApplication.WorkBooks[1].Worksheets['Alternatives'];
  Sheet2 := ExcelApplication.WorkBooks[1].Worksheets['Plans'];
  }
  grdmatrizconseq.Columns.Clear;
  grdmatrizconseq.ClearCells;
  grdmatrizconseq.TotalRows := 10 + Length(Populacao);
  grdmatrizconseq.columns.Add;
  for I := 0 to Length(ObjSelecionados) - 1 do
  begin
  grdmatrizconseq.columns.Add;
  grdmatrizconseq.Cells[i + 1,0] := ObjSelecionados[i];
  //sheet1.cells[1,2+i] := ObjSelecionados[i];
  if i > 0 then
  begin
  grdmatrizconseq.Cells[i + 1,1] := '1';
  ///sheet1.cells[2,2+i] := '1';
  end
  else
  begin
  grdmatrizconseq.Cells[i + 1,1] := '0';
  //sheet1.Cells[2,i + 2] := '0'
  end;
  end;
  grdmatrizconseq.Cells[0,0] := 'Criteria:';
  grdmatrizconseq.Cells[0,1] := 'Preference Direction:';
  grdmatrizconseq.Cells[0,6] := 'Number of Levels:';
  grdmatrizconseq.Cells[0,7] := 'Alternatives:';
  grdmatrizconseq.Cells[1,7] := 'Consequence Matrix:';


  {
  sheet1.Cells[1,1] := 'Criteria:';
  sheet1.Cells[2,1] := 'Preference Direction:';
  sheet1.Cells[7,1] := 'Number of Levels:';
  sheet1.Cells[8,1] := 'Alternatives:';
  sheet1.Cells[8,2] := 'Consequence Matrix:';
   }

  //** Adicionar apenas alternativas não repetidas **//
  for i := 0 to Length(Populacao) - 1 do
  begin
  grdmatrizconseq.Cells[0,8+i] := 'Plan ' + inttostr(i+1);
  //sheet1.Cells[9+i,1] := 'Plan ' + inttostr(i+1);
    for t := 0 to Length(Matriz) - 1 do
    begin
    counter1 := 0;
      for j := 0 to Length(PerformancePopulacao[i]) - 1 do

        if PerformancePopulacao[i, j] = Matriz[t, j] then
          Inc(counter1);
    end;

    if counter1 = Length(PerformancePopulacao[i]) then
      continue
    else
    begin
      Setlength(Matriz, Length(Matriz) + 1, Length(PerformancePopulacao[i])-1);
      for j := 0 to Length(PerformancePopulacao[i]) - 2 do
      begin
        Matriz[Length(Matriz) - 1, j] := PerformancePopulacao[i, j];
        grdmatrizconseq.Cells[j + 1,Length(Matriz)+7] :=  floatToStr(TIWUserSession(WebApplication.Data).PerformancePopulacao[i, j]);
        //sheet1.Cells[Length(Matriz)+8,j+2] :=  floatToStr(PerformancePopulacao[i, j]);
      end;


    
      end; //** else
  end;

  {if SaveDialog1.Execute then
  begin
  ExcelApplication.WorkBooks[1].SaveAs(SaveDialog1.FileName + '.xls');
  end;
  ExcelApplication.WorkBooks[1].Close(SaveChanges := false);

  ExcelApplication := Unassigned; }

  TIWAdvWebGridExcelIO1.AdvWebGrid:= grdmatrizconseq;
 TIWAdvWebGridExcelIO1.XLSExport('Consequence Matrix - ' + problema + '.xls', WebApplication);
 //Refresh;

 IwTimer1.Enabled := False;
  //TIWAdvWebGridExcelIO1.Free;

  end;
end;


Procedure  TIWForm2.SalvarBD;
var
i, j : integer;
Id_Problema, prefdir : string;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  if retomada = false then
  begin
  with UserSession.ZQuery1 do
  begin
  sql.Text := 'INSERT INTO problem (User, Name, Creation_Date, Conclusion_Date, Context) VALUES ("' + Id_User + '", "' + Problema + '", "' + datainicio + '", "' + datetostr(now) + '", "' + TiwUserSession(WebApplication.Data).ContextProblem + '")';
  ExecSQL;

  SQL.Text := 'SELECT Id_Problem FROM problem WHERE User ="' + (Id_User)
        + '"';
      ExecSQL;
      open;
      Last;
      Id_Problema := FieldByName('Id_Problem').AsString;
      close;

  for i := 0 to Length(TiwUserSession(WebApplication.Data).Actor) - 1 do
  begin
  sql.Text := 'INSERT INTO actors (Id_Problem, ActorType, ActorDescription) VALUES ("' + Id_Problema + '", "' + Actor[i] + '", "' + ActorDescription[i] + '")';
  ExecSQL;
  end;


  sql.Text := 'INSERT INTO restrictions (Id_Problem, Total_Period, Regular_Production, Overtime_Production, Subcontrating_Production, Stock, Ini_Employees, Max_Employees) VALUES ("' + Id_Problema + '" ,"' + inttostr(Periodo) + '","' + floattostr(TIWUserSession(WebApplication.Data).limites[0]) + '", "' + floattostr(limites[1]) + '", "' + floattostr(limites[2]) + '", "' + floattostr(limites[3]) + '", "' + inttostr(NInicialFunc) + '", "' + floattostr(limites[4]) + '" )';
  ExecSQL;

  for i := 0 to Periodo - 1 do
  begin
   sql.Text := 'INSERT INTO demand (Id_Problem, Total_Period, Period, Demand) VALUES ("' + Id_Problema + '" ,"' + inttostr(Periodo) + '","' + inttostr(i+1) + '","' + floattostr(TIWUserSession(WebApplication.Data).demanda[i]) + '")';
  ExecSQL;
  end;

  for i := 0 to length(ObjSelecionados) - 1 do
  begin
  if TIWUserSession(WebApplication.Data).ObjSelecionados[TIWUserSession(WebApplication.Data).ordemsel[i]] = 'Cost'  then
  begin
    sql.Text := 'INSERT INTO obj_functions (Id_Problem, Objective, Pref_dir) VALUES ("' + Id_Problema + '", "' + TIWUserSession(WebApplication.Data).ObjSelecionados[TIWUserSession(WebApplication.Data).ordemsel[i]] + '","' + '0' + '")';
    ExecSQL;

      // salvando no BD
      SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(custos[0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(custos[1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(custos[2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(custos[3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(custos[4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(custos[5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end
  else
  begin
    sql.Text := 'INSERT INTO obj_functions (Id_Problem, Objective, Pref_dir) VALUES ("' + Id_Problema + '", "' + ObjSelecionados[ordemsel[i]] + '","' + '1' + '")';
    ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(TIWUserSession(WebApplication.Data).ParametrosFuncaoObjetivo[TIWUserSession(WebApplication.Data).ordemsel[i]-1, 0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

          SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end;
  end;
  end;
  end // if retomada = false
  else
  begin
    with UserSession.ZQuery1 do
  begin
  Id_problema := Id_Problem;
  SQL.Text := 'UPDATE problem SET Name ="' + Problema + '", Creation_Date ="' + datainicio + '", Conclusion_Date ="' + datetostr(now) + '"    WHERE Id_Problem = "' + Id_Problema
        + '";';
      ExecSQL;

  SQL.Text := 'UPDATE restrictions SET Total_Period ="' + inttostr(Periodo) + '", Regular_Production ="' + floattostr(TIWUserSession(WebApplication.Data).limites[0]) + '",Overtime_Production ="' + floattostr(limites[1]) + '", Subcontrating_Production ="' + floattostr(limites[2]) + '",  Stock ="' + floattostr(limites[3]) + '", Ini_Employees ="' + inttostr(NInicialFunc) + '", Max_Employees ="' + floattostr(limites[4]) + '"  WHERE Id_Problem = "' + Id_Problema
        + '";';
      ExecSQL;

  for i := 0 to Periodo - 1 do
  begin
   SQL.Text := 'UPDATE demand SET Total_Period ="' + inttostr(Periodo) + '", Demand = "' + FloatToStr(demanda[i]) + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Period = "' + inttostr(i+1) + '";';
      ExecSQL;
  end;

  for i := 0 to length(ObjSelecionados) - 1 do
  begin
  if TIWUserSession(WebApplication.Data).ObjSelecionados[TIWUserSession(WebApplication.Data).ordemsel[i]] = 'Cost'  then
  begin


      // salvando no BD
      SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(custos[0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(custos[1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(custos[2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(custos[3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(custos[4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(custos[5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end
  else
  begin


    SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(TIWUserSession(WebApplication.Data).ParametrosFuncaoObjetivo[TIWUserSession(WebApplication.Data).ordemsel[i]-1, 0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

          SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end;
  end;
  end;
  end;





  end;
end;



procedure TIWForm2.btnImportarClick(Sender: TObject);
var
i, j, auxcount : integer;
user_ ,datafinal: string;
begin
LimpaVariaveis;
With TIWUserSession(WebApplication.Data) do
begin
lnkplanos.Visible := False;
  with UserSession.ZQuery1 do
  begin
    // inicialmente contar quantos problemas o mesmo usuario já cadastrou
    SQL.Text := 'SELECT COUNT(name) FROM problem WHERE user = "' +
      (ID_user) + '";';
    ExecSQL;
    open;
    auxcount := FieldByName('COUNT(name)').AsInteger;
    close;

    { após saber quantos problemas o mesmo usuario cadastrou é possivel fazer um
    laço de repetição variando de zero ate essa quantidade-1 para pegar os nomes dos problemas }
    // VERIFICAR COMO FAZER ESSE PROCEDIMENTO

    i := 0;
    j := 0;
    SQL.Text := 'SELECT * FROM problem';
    ExecSQL;
    open;
    while i < auxcount do
    begin
      user_ := FieldByName('user').AsString;
      datafinal := FieldByName('conclusion_date').AsString;

      if ID_user = user_ then
      begin
        Setlength(nameproblem, j + 1);
        Setlength(idproblemasuser,j + 1);
        Setlength(datainicioproblema, j + 1);
        Setlength(datafinalproblema, j + 1);

        nameproblem[j] :=FieldByName('name').AsString;
        idproblemasuser[j] := FieldByName('id_problem').AsString;
        datainicioproblema[j] := FieldByName('creation_date').AsString;
        datafinalproblema[j] := FieldByName('conclusion_date').AsString;
        if datafinalproblema[j] = '' then
        datafinalproblema[j] := '02/01/1900';

        INC(j);
        // end;
        INC(i);
      end;
      next;
    end;
    close;
  end;

  if length(nameproblem) = 0 then
  begin
    WebApplication.ShowMessage('There is no problem to be continued');
    Exit;
  end;

  rgnopcaoinput.Visible := false;
  rgnretomada.Visible := true;
  btnauxchoose.Enabled := true;
  Retomada := True;

  // colocar o vetor do nome no listbox
  lstbxProjects.Items.Clear;
  for i := 0 to length(nameproblem) - 1 do
  begin
    lstbxProjects.Items.Add(nameproblem[i] + ' - (' + datainicioproblema[i] + ')');
  end;
end;
end;

procedure TIWForm2.btnImportClick(Sender: TObject);
var
i,j : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
with UserSession.ZQuery1 do
  begin

  SQL.Text := 'SELECT * FROM prob_padrao';
    ExecSQL;
    open;
    idproblemapadrao := FieldByName('Id_problem').Asstring;
    close;

   SetLength(Limites, 5);
   Setlength(Custos, 6);

  SQL.Text := 'SELECT * FROM restrictions WHERE Id_Problem ="' + idproblemapadrao
        + '"';
      ExecSQL;
      open;
      Periodo := strtoint(FieldByName('Total_Period').AsString);
      Limites[0] := strtofloat(FieldByName('Regular_Production').AsString);
      Limites[1] := strtofloat(FieldByName('Overtime_Production').AsString);
      Limites[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
      Limites[3] := strtofloat(FieldByName('Stock').AsString);
      Limites[4] := strtofloat(FieldByName('Max_Employees').AsString);
      NinicialFunc := strtoint(FieldByName('Ini_Employees').AsString);
      close;


    SetLength(Demanda, Periodo);
    for i := 0 to Periodo - 1 do
    begin
    SQL.Text := 'SELECT * FROM demand WHERE Id_Problem ="' + idproblemapadrao
        + '" AND Period ="' + inttostr(i+1) + '"';
      ExecSQL;
      open;
      Demanda[i] := strtofloat(FieldByName('Demand').AsString);
      close;
    end;






      SetLength(Custos, 6);


      SQL.Text := 'SELECT * FROM obj_functions WHERE Id_Problem ="' + idproblemapadrao
        +'"';
      ExecSQL;
      open;
      First;
      i := 0;
      j := 0;
      while not eof do
      begin
      SetLength(ObjSelecionados, i+1);
      TIWUserSession(WebApplication.Data).ObjSelecionados[i] := FieldByName('Objective').AsString;
      if ObjSelecionados[i] = 'Cost' then
      begin
        TIWUserSession(WebApplication.Data).Custos[0] := strtofloat(FieldByName('Regular_Production').AsString);
        Custos[1] := strtofloat(FieldByName('Overtime_Production').AsString);
        Custos[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
        Custos[3] := strtofloat(FieldByName('Stock').AsString);
        Custos[4] := strtofloat(FieldByName('Hiring').AsString);
        Custos[5] := strtofloat(FieldByName('Resignation').AsString);
      end
      else
      begin
       SetLength(ParametrosFuncaoObjetivo, j+1, 6);
       TIWUserSession(WebApplication.Data).ParametrosFuncaoObjetivo[j,0] := strtofloat(FieldByName('Regular_Production').AsString);
       ParametrosFuncaoObjetivo[j,1] := strtofloat(FieldByName('Overtime_Production').AsString);
       ParametrosFuncaoObjetivo[j,2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
       ParametrosFuncaoObjetivo[j,3] := strtofloat(FieldByName('Stock').AsString);
       ParametrosFuncaoObjetivo[j,4] := strtofloat(FieldByName('Hiring').AsString);
       ParametrosFuncaoObjetivo[j,5] := strtofloat(FieldByName('Resignation').AsString);
       inc(j);
      end;
      inc(i);
      Next;
      end;
      Close;

      rgninput.Visible := True;
      rgnopcaoinput.Visible := False;

      if Periodo = 12 then
      rgrpPeriodo.ItemIndex := 0;
      if Periodo = 6 then
      rgrpPeriodo.ItemIndex := 1;
      if Periodo = 4 then
      rgrpPeriodo.ItemIndex := 2;

      edtnomeproblema.Text := 'Default Problem';

      if (Limites[0] = 0) or (Custos[0] = 0) then
      begin
        cbxproducaoregular.Checked := False;
        cbxproducaoregularClick(Self);
      end;

      if (Limites[1] = 0) or (Custos[1] = 0) then
      begin
        cbxHorasExtras.Checked := False;
        cbxHorasExtrasClick(Self);
      end;

      if (Limites[2] = 0) or (Custos[2] = 0) then
      begin
        cbxSubcontratacao.Checked := False;
        cbxSubcontratacaoClick(Self);
      end;

       if (Limites[3] = 0) or (Custos[3] = 0) then
      begin
        cbxEstoque.Checked := False;
        cbxEstoqueClick(Self);
      end;

      if (Custos[4] = 0) then
      begin
        cbxcontrat.Checked := False;
        cbxcontratClick(Self);
      end;

      if (Custos[5] = 0) then
      begin
        cbxdem.Checked := False;
        cbxdemclick(Self);
      end;


      edtprodregular.Text := FloatToStr(limites[0]);
      edthorasextras.Text := FloatToStr(limites[1]);
      edtsubcontrat.Text := FloatToStr(limites[2]);
      edtestoq.Text := FloatToStr(limites[3]);
      edtfunc.Text := FloatToStr(limites[4]);
      edtinifunc.Text := IntToStr(NinicialFunc);
      edtprodregular.Text := trocavp(edtprodregular.Text);
      edthorasextras.Text := trocavp(edthorasextras.Text);
      edtsubcontrat.Text := trocavp(edtsubcontrat.Text);
      edtestoq.Text := trocavp(edtestoq.Text);
      edtfunc.Text := trocavp(edtfunc.Text);
      edtinifunc.Text := trocavp(edtinifunc.Text);
      btnsalvaretapa1Click(Self);

      lstbxperformanceobj.Items.Clear;
      lstbxperformanceobj.Items.Add('Selected Objectives:');
      for j := 0 to length(ObjSelecionados) - 1 do
      begin
        lstbxperformanceobj.Items.Add(ObjSelecionados[j]);
        for i := 0 to cmbbxperformanceobj.Items.Count - 1 do
        begin
          cmbbxperformanceobj.ItemIndex := i;
          if cmbbxperformanceobj.Text = objSelecionados[j] then
          begin
            cmbbxperformanceobj.Items.Delete(cmbbxperformanceobj.ItemIndex);
          end;
        end;
      end;
      problemapadrao := true;
      btnavancaretapa1.Enabled := True;
      datainicio := datetostr(now);
  end;

  end;
end;


procedure TIWForm2.btnInputManualClick(Sender: TObject);
begin
With TIWUserSession(WebApplication.Data) do
begin
LimpaVariaveis;
rgninput.Visible := True;
rgnopcaoinput.Visible := False;
Setlength(ObjSelecionados, 1);
ObjSelecionados[0] := 'Cost';
datainicio := datetostr(now);
retomada := False;
lnkplanos.Visible := False;
end;
end;




procedure TIWForm2.btnLoginClick(Sender: TObject);
begin
rgnintro.Visible := false;
rgnlogin.Visible := true;
end;

procedure TIWForm2.btnNewActorRoleClick(Sender: TObject);
begin
if (edtnewactor.Text <> '') and (edtnewactor.Text <> '[In case the actor role is not included in the list]') then
cmbxactor.Items.Add(edtnewactor.Text);
edtnewactor.text := '[In case the actor role is not included in the list]';
end;

procedure TIWForm2.btnobjClick(Sender: TObject);
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    if cmbbxperformanceobj.Text = '' then
      WebApplication.ShowMessage('Select the objective correctly!')
    else
    begin
      lstbxperformanceobj.Items.Add(cmbbxperformanceobj.Text);
      SetLength(ObjSelecionados, Length(ObjSelecionados)+1);
      ObjSelecionados[Length(ObjSelecionados)-1] := cmbbxperformanceobj.Text;
      cmbbxperformanceobj.Items.Delete(cmbbxperformanceobj.ItemIndex);
      if (rgrpPeriodo.ItemIndex <> -1) and (btnsalvaretapa1.Enabled = False) then
        btnavancaretapa1.Enabled := True;

      cmbbxperformanceobj.ItemIndex := -1;
    end;
  end;
end;

procedure TIWForm2.btnperformanceobjClick(Sender: TObject);
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    lstbxperformanceobj.Items.Add(edtotherperformanceobj.Text);
    SetLength(ObjSelecionados, Length(ObjSelecionados)+1);
    ObjSelecionados[Length(ObjSelecionados)-1] := edtotherperformanceobj.Text;
    edtotherperformanceobj.Text := '';
    if (rgrpPeriodo.ItemIndex <> -1) and (btnsalvaretapa1.Enabled = False) then
      btnavancaretapa1.Enabled := True;
  end;
end;

procedure TIWForm2.btnredefiniretapa1Click(Sender: TObject);
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  edtprodregular.Enabled := True;
  edthorasextras.Enabled := True;
  edtsubcontrat.Enabled := True;
  edtestoq.Enabled := True;
  edtfunc.Enabled := True;
  edtinifunc.Enabled := True;

  edtprodregular.Clear;
  edthorasextras.Clear;
  edtsubcontrat.Clear;
  edtestoq.Clear;
  edtfunc.Clear;
  edtinifunc.Clear;


  cbxproducaoregular.checked := True;
  cbxHorasExtras.checked := True;
  cbxSubcontratacao.checked := True;
  cbxEstoque.checked := True;
  cbxcontrat.checked := True;
  cbxdem.checked := True;
 // edtnomeproblema.Text := '';
  rgrpPeriodo.ItemIndex := -1;

  //btnavancaretapa1.Enabled := False;
  //btnsalvaretapa1.enabled := True;
 // retomada := False;
  end;
end;

procedure TIWForm2.btnredefiniretapa2Click(Sender: TObject);
begin
With TIWUserSession(WebApplication.Data) do
begin
Chrtdemanda.Series[0].Clear;
Chrtdemanda.Series[1].Clear;
Chrtdemanda.Series[0].Visible := False;
Chrtdemanda.Series[1].Visible := False;
Chrtdemanda.Series[2].Visible := True;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));
Finalize(Demanda);
btnsalvaretapa2.Enabled := True;
btnredefiniretapa2.Enabled := False;
btnavancaretapa2.Enabled := False;
//retomada := False;

end;
end;

procedure TIWForm2.btnredefiniretapa3Click(Sender: TObject);
begin
 cmbbxobjetivos.Enabled := True;
 grdfuncaoobjetivo.Enabled := True;
 //btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 if lstbxFuncoesObjetivo.Items.Count > 0 then
 lstbxFuncoesObjetivo.Items.Delete(lstbxFuncoesObjetivo.Items.Count-1);
end;

procedure TIWForm2.btnredefinirnovoproblemaClick(Sender: TObject);
begin
  With TIWUserSession(WebApplication.Data) do
  begin


  finalize(actor);
  finalize(actordescription);
  edtdecisionproblem.Text := '';
  memoContext.Text := '';

  lstbxperformanceobj.Items.Clear;
  lstbxperformanceobj.Items.Add('Selected Objectives:');
  lstbxperformanceobj.Items.Add('Cost');
  cmbbxperformanceobj.Items.Clear;
  cmbbxperformanceobj.Items.Add('Quality');
  cmbbxperformanceobj.Items.Add('Reliability');
  cmbbxperformanceobj.Items.Add('Speed');
  cmbbxperformanceobj.Items.Add('Flexibility');
  cmbbxperformanceobj.ItemIndex := -1;

  cmbxactor.Items.Clear;
  cmbxactor.items.Add('Decision Maker');
  cmbxactor.items.Add('Analyst');
  cmbxactor.items.Add('Facilitator');
  cmbxactor.items.Add('Client');
  cmbxactor.items.Add('Expert/Specialist');
  cmbxactor.items.Add('Stakeholder');
  edtnewactor.text := '[In case the actor role is not included in the list]';
  edtactordescription.text := '';
  grdactorroles.ClearCells;
  grdactorroles.RowCount := 0;
  grdactorroles.totalrows := 0;
  cmbxactor.ItemIndex := -1;
  edtotherperformanceobj.Clear;
  Finalize(ObjSelecionados);
  Setlength(ObjSelecionados, 1);
  ObjSelecionados[0] := 'Cost';
  rgnotherperformanceobj.Visible := False;
  //btnavancaretapa1.Enabled := False;
  //btnsalvaretapa1.enabled := True;
 // retomada := False;
  end;
end;

procedure TIWForm2.btnRegisterClick(Sender: TObject);
begin
 iwrgnCaixaDialogoRU.Visible := true;
end;

procedure TIWForm2.btnRegisternewClick(Sender: TObject);
begin
With TIWUserSession(WebApplication.Data) do
begin
LimpaVariaveis;
//rgninput.Visible := True;
rgnnovoproblema.Visible := true;
rgnopcaoinput.Visible := False;
grdactorroles.rowcount := 0;
grdactorroles.totalrows := 0;
Setlength(ObjSelecionados, 1);
btnredefinirnovoproblemaClick(self);
ObjSelecionados[0] := 'Cost';
datainicio := datetostr(now);
retomada := False;
problemapadrao := false;
lnkplanos.Visible := False;
end;
end;

procedure TIWForm2.btnreiniciaretapa3Click(Sender: TObject);
var
i : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
 Finalize(custos);
 Finalize(ParametrosFuncaoObjetivo);
 Finalize(ordemsel);
 auxiliar := 0;
 cmbbxObjetivos.Items.Clear;
 for I := 0 to Length(ObjSelecionados)-1 do
 cmbbxobjetivos.Items.Add(objselecionados[i]);
 cmbbxobjetivos.ItemIndex := 0;
 cmbbxobjetivos.Enabled := True;
 //grdfuncaoobjetivo.Enabled := True;
 //btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 btnexecutar.Enabled := False;
 lstbxFuncoesObjetivo.Clear;
 //Retomada := False;
end;
end;

procedure TIWForm2.btnsalvaretapa1Click(Sender: TObject);
var
number : double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  edtprodregular.Text := trocavp(edtprodregular.Text);
  edthorasextras.Text := trocavp(edthorasextras.Text);
  edtsubcontrat.Text := trocavp(edtsubcontrat.Text);
  edtestoq.Text := trocavp(edtestoq.Text);
  edtfunc.Text := trocavp(edtfunc.Text);
  edtinifunc.Text := trocavp(edtinifunc.Text);
  if (TryStrToFloat(edtprodregular.Text, number) = True) and
   (TryStrToFloat(edthorasextras.Text, number) = True) and
   (TryStrToFloat(edtsubcontrat.Text, number) = True) and
   (TryStrToFloat(edtestoq.Text, number) = True) and
   (TryStrToFloat(edtfunc.Text, number) = True) and
   (TryStrToFloat(edtinifunc.Text, number) = True) then
  begin
  Setlength(Limites, 5);

  Limites[0] := Strtofloat(edtprodregular.Text); // ** Max prod regular
  Limites[1] := Strtofloat(edthorasextras.Text); // ** Max prod h extra
  Limites[2] := Strtofloat(edtsubcontrat.Text); // ** Maxima prod subcontratada
  Limites[3] := Strtofloat(edtestoq.Text); // **Máximo estoque por período
  Limites[4] := Strtofloat(edtfunc.Text); // ** Número máximo de funcionários
  NinicialFunc := Strtoint(edtinifunc.Text); // ** Número inicial de funcionários
  edtprodregular.Enabled := False;
  edthorasextras.Enabled := False;
  edtsubcontrat.Enabled := False;
  edtestoq.Enabled := False;
  edtfunc.Enabled := False;
  edtinifunc.Enabled := False;

  btnsalvaretapa1.enabled := False;
  if (lstbxperformanceobj.Items.Count > 2) and (rgrpPeriodo.ItemIndex <> -1) then
  btnavancaretapa1.Enabled := True;
  end
  else
  begin
  WebApplication.ShowMessage('Fill in the information correctly!');
  exit;
  end;
  end;
end;



procedure TIWForm2.btnsalvaretapa2Click(Sender: TObject);
var
i : integer;
number, aux : double;
nome : string;
begin
With TIWUserSession(WebApplication.Data) do
begin
SetLength(Demanda, Periodo);
for i := 0 to periodo - 1 do
begin
  nome := grddemanda.DynEdits[1,i];
  nome := trocavp(nome);
  if (TryStrToFloat(nome, number) = True) then
  begin
  grddemanda.DynEdits[1,i] := trocavp(grddemanda.DynEdits[1,i]);
  demanda[i] := StrToFloat(grddemanda.DynEdits[1,i]);
  aux := aux + demanda[i];
  end
  else
  begin
  WebApplication.ShowMessage('Fill in the information correctly!');
  exit;
  end;
end;
aux := aux / periodo;
DemandaMedia := aux;
Chrtdemanda.Series[0].Visible := True;
Chrtdemanda.Series[1].Visible := True;
for i := 0 to Periodo -1  do
begin
chrtDemanda.Series[0].AddXY(i+1, demanda[i]);
chrtDemanda.Series[1].AddXY(i+1, aux);
end;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));

btnsalvaretapa2.Enabled := False;
btnredefiniretapa2.Enabled := true;
btnavancaretapa2.Enabled := True;

end;
end;



procedure TIWForm2.btnsalvaretapa3Click(Sender: TObject);
var
i, aux : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
SetLength(Custos, 6);
SetLength(ParametrosFuncaoObjetivo, length(ObjSelecionados)-1, 6);

SetLength(ordemsel, length(ObjSelecionados));


aux := 0;

if cmbbxobjetivos.Text = 'Cost'  then
begin
 for I := 0 to 5 do
 begin
 grdfuncaoobjetivo.DynEdits[i,0] := trocavp(grdfuncaoobjetivo.DynEdits[i,0]);
 custos[i] := StrToFloat(grdfuncaoobjetivo.DynEdits[i,0]);
 end;
 lstbxFuncoesObjetivo.Items.Add('Min. ' + cmbbxObjetivos.Text + ' = ' + floattostr(custos[0]) +
 '*x1 + '+ floattostr(custos[1]) +'*x2 + '+ floattostr(custos[2])+ '*x3 + '+ floattostr(custos[3])+ '*x4 + '
 + floattostr(custos[4])+ '*x5 + ' + floattostr(custos[5])+ '*x6');
end
else
begin
  for I := 0 to length(objselecionados) - 1 do
  begin
    if objselecionados[i] = cmbbxobjetivos.Text then
    aux := i;
  end;
 for I := 0 to 5 do
 begin
 grdfuncaoobjetivo.DynEdits[i,0] := trocavp(grdfuncaoobjetivo.DynEdits[i,0]);
 ParametrosFuncaoObjetivo[aux-1, i] := StrToFloat(grdfuncaoobjetivo.DynEdits[i,0]);
 end;
 lstbxFuncoesObjetivo.Items.Add('Max. ' + cmbbxObjetivos.Text + ' = ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 0]) +
 '*x1 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 1]) +'*x2 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 2])+ '*x3 '
 + floattostr(ParametrosFuncaoObjetivo[aux-1, 3]) +'*x4 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 4]) +'*x5 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 5]) +'*x6');
 end;


 ordemsel[auxiliar] := aux;
 inc(auxiliar);

 cmbbxobjetivos.Items.Delete(cmbbxobjetivos.ItemIndex);
 cmbbxobjetivos.Enabled := True;
 grdfuncaoobjetivo.Enabled := True;
 //btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 if cmbbxobjetivos.Items.Count = 0 then
 begin
 cmbbxobjetivos.Enabled := False;
 grdfuncaoobjetivo.Enabled := False;
 btnsalvaretapa3.Enabled := False;
 //btnredefiniretapa3.Enabled := True;
 btnexecutar.Enabled := True;
 end;
end;
end;



procedure TIWForm2.btnSaveProblemClick(Sender: TObject);
begin

 if edtDecisionProblem.Text = '' then
  begin
  WebApplication.ShowMessage('Fill in the problem name to continue!');
  exit;
  end;


if (lstbxperformanceobj.Items.Count <= 2) then
  begin
  WebApplication.ShowMessage('Define the objectives to continue!');
  exit;
  end;


  if (grdactorroles.Rowcount = 0) then
  begin
  WebApplication.ShowMessage('Define an actor to continue!');
  exit;
  end;



edtnomeproblema.Text := edtDecisionProblem.Text;
TIWUserSession(WebApplication.Data).ContextProblem := memoContext.Text;
rgninput.Visible := true;
rgnnovoproblema.Visible := false;
end;

procedure TIWForm2.cbxcontratClick(Sender: TObject);
begin
if cbxcontrat.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,4] := '0';
grdfuncaoobjetivo.Columns.Items[4].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[4].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxdemClick(Sender: TObject);
begin
if cbxdem.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,5] := '0';
grdfuncaoobjetivo.Columns.Items[5].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[5].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxEstoqueClick(Sender: TObject);
begin
if cbxestoque.Checked = False then
begin
edtestoq.Text := '0';
edtestoq.Enabled := False;
end
else
edtestoq.Enabled := True;

if cbxestoque.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,3] := '0';
grdfuncaoobjetivo.Columns.Items[3].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[3].ColumnType := ctDynEdit;

end;



procedure TIWForm2.cbxHorasExtrasClick(Sender: TObject);
begin
if cbxhorasextras.Checked = False then
begin
edthorasextras.Text := '0';
edthorasextras.Enabled := False;
end
else
edthorasextras.Enabled := True;

if cbxhorasextras.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,1] := '0';
grdfuncaoobjetivo.Columns.Items[1].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[1].ColumnType := ctDynEdit;

end;


procedure TIWForm2.cbxproducaoregularClick(Sender: TObject);
begin
if cbxproducaoregular.Checked = False then
begin
edtprodregular.Text := '0';
edtprodregular.Enabled := False;
end
else
edtprodregular.Enabled := True;

if cbxproducaoregular.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,0] := '0';
grdfuncaoobjetivo.Columns.Items[0].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[0].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxSubcontratacaoClick(Sender: TObject);
begin
if cbxsubcontratacao.Checked = False then
begin
edtsubcontrat.Text := '0';
edtsubcontrat.Enabled := False;
end
else
edtsubcontrat.Enabled := True;

if cbxsubcontratacao.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,2] := '0';
grdfuncaoobjetivo.Columns.Items[2].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[2].ColumnType := ctDynEdit;
end;








procedure TIWForm2.edtprodregularAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtprodregular.Text[Length(edtprodregular.Text)] <> '0') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '1') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '2') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '3') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '4') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '5') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '6') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '7') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '8') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '9') then
    edtprodregular.Text := Copy(edtprodregular.Text, 1, length(edtprodregular.Text) - 1)
    else
    edtprodregular.Text := edtprodregular.Text;
end;






procedure TIWForm2.edtsubcontratAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '0') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '1') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '2') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '3') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '4') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '5') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '6') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '7') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '8') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '9') then
    edtsubcontrat.Text := Copy(edtsubcontrat.Text, 1, length(edtsubcontrat.Text) - 1)
    else
    edtsubcontrat.Text := edtsubcontrat.Text;
end;



procedure TIWForm2.IWAppFormCreate(Sender: TObject);
begin
if TIWUserSession(WebApplication.Data).reinicia = 1 then
begin
rgnopcaoinput.visible := True;
rgnintro.visible := False;
TIWUserSession(WebApplication.Data).reinicia := 0;
end;
if TIWUserSession(WebApplication.Data).reinicia = 2 then
begin
rgnfuncaoobjetivo.visible := True;
rgnintro.visible := False;
TIWUserSession(WebApplication.Data).reinicia := 0;
end;
end;

procedure TIWForm2.IWButton1Click(Sender: TObject);
begin
With TIWUserSession(WebApplication.Data) do
begin
if cmbbxperformanceobj.Text = '' then
WebApplication.ShowMessage('Select the objective correctly!')
else
begin
lstbxperformanceobj.Items.Add(cmbbxperformanceobj.Text);
SetLength(ObjSelecionados, Length(ObjSelecionados)+1);
ObjSelecionados[Length(ObjSelecionados)-1] := cmbbxperformanceobj.Text;
cmbbxperformanceobj.Items.Delete(cmbbxperformanceobj.ItemIndex);
if (rgrpPeriodo.ItemIndex <> -1) and (btnsalvaretapa1.Enabled = False) then
btnavancaretapa1.Enabled := True;
end;
end;
end;



procedure TIWForm2.IWButton4Click(Sender: TObject);
begin
  rgnlogin.Visible := false;
  rgnintro.Visible := true;
  lblwarningemail.Visible := false;
  lblwarningpassword.Visible := false;

  /// /////////////////////////////////////////////////////////////////////////////
  /// ////////////* Limpando edit's de login, validacao *//////////////////////////
  edtemail.Clear;
  edtSenha.Clear;
  /// /////////////////////////////////////////////////////////////////////////////
end;

procedure TIWForm2.lnkplanosClick(Sender: TObject);
var
i, counter2, k, z : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
    grdplanos.columns.Clear;
    counter2 := -1;
    grdplanos.TotalRows := 3;
    for I := 0 to 8 do
    begin
    grdplanos.columns.Add;
    end;
    for I := 0 to Length(Populacao) - 1 do
    begin
    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
    grdplanos.Cells[0,Counter2+1] := 'Plan ' + inttostr(i+1) + ':';
    //sheet2.Cells[Counter2+2, 1] := 'Plan ' + inttostr(i+1) + ':';
    for k := 0 to Length(Populacao[i]) - 1 do
    begin
    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
    //grdplanos.Cells[0,Counter2+1] := inttostr(z+1);
      for z := 0 to Length(Populacao[i, k]) - 1 do
        begin
       grdplanos.Cells[Z + 1,Counter2+1] := floatToStr(Populacao[i, k, z]);
       //sheet2.Cells[Counter2+2,z+2] := floatToStr(Populacao[i, k, z]);
        end;
    end;
    end;
    grdplanos.Cells[0,0] := 'Plan';
   grdplanos.Cells[1,0] := 'Regular Production';
   grdplanos.Cells[2,0] := 'Overtime Production';
   grdplanos.Cells[3,0] := 'Subcontrating Production';
   grdplanos.Cells[4,0] := 'Stock';
   grdplanos.Cells[5,0] := 'Hiring';
   grdplanos.Cells[6,0] := 'Resignation';
   grdplanos.Cells[7,0] := 'Total Employees';
TIWAdvWebGridExcelIO2.AdvWebGrid := grdplanos;
  TIWAdvWebGridExcelIO2.XLSExport('Plans_' + problema + '.xls', WebApplication);
  //Inherited;
end;
end;

procedure TIWForm2.IWLink11Click(Sender: TObject);
begin
rgnopcaoinput.Visible := true;
rgnnovoproblema.Visible := false;
end;

procedure TIWForm2.IWLink13Click(Sender: TObject);
begin
rgnotherperformanceobj.Visible := true;
end;

procedure TIWForm2.IWLink15Click(Sender: TObject);
begin
edtotherperformanceobj.Text := '';
rgnotherperformanceobj.Visible := false;
end;



procedure TIWForm2.IWLink2Click(Sender: TObject);
begin
btnredefiniretapa3Click(Self);
btnredefiniretapa2Click(Self);
btnredefiniretapa1Click(Self);
rgninput.Visible := False;
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := False;
rgnintro.Visible := False;
rgnlogin.Visible := False;
rgnretomada.Visible := False;
rgnopcaoinput.Visible := True;
end;





procedure TIWForm2.IWLink4Click(Sender: TObject);
begin
rgnopcaoinput.Visible := True;
rgnretomada.Visible := False;
end;

procedure TIWForm2.IWLink5Click(Sender: TObject);
begin
rgnopcaoinput.visible := False;
rgnlogin.Visible := True;
edtemail.Clear;
edtSenha.Clear;
end;

procedure TIWForm2.IWLink6Click(Sender: TObject);
begin
btnredefiniretapa3Click(Self);
btnredefiniretapa2Click(Self);
btnredefiniretapa1Click(Self);
rgninput.Visible := False;
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := False;
rgnintro.Visible := False;
rgnlogin.Visible := False;
rgnretomada.Visible := False;
rgnopcaoinput.Visible := True;
end;

procedure TIWForm2.IWLink7Click(Sender: TObject);
begin
rgndemanda.Visible := False;
rgninput.Visible := True;
btnredefiniretapa2Click(Self);
end;

procedure TIWForm2.IWLink8Click(Sender: TObject);
begin
rgnfuncaoobjetivo.Visible := False;
rgndemanda.Visible := True;
btnredefiniretapa3Click(Self);
end;

procedure TIWForm2.IWTimer1Timer(Sender: TObject);
begin
if IwTimer1.Enabled = True then
begin
  Exportar;
end

end;

procedure TIWForm2.lnkesqueceusenhaClick(Sender: TObject);
begin
iwrgnCaixaDialogoRU2.Visible := true;
end;

procedure TIWForm2.rgndemandaCreate(Sender: TObject);
begin
rgndemanda.align := AlClient;
end;

procedure TIWForm2.rgnfuncaoobjetivoCreate(Sender: TObject);
begin
rgnfuncaoobjetivo.align := AlClient;
end;

procedure TIWForm2.rgninputCreate(Sender: TObject);
begin
rgninput.align := AlClient;
end;

procedure TIWForm2.rgnprobCreate(Sender: TObject);
begin
rgnopcaoinput.Align := AlClient;
end;

procedure TIWForm2.rgnretomadaCreate(Sender: TObject);
begin
rgnretomada.Align := AlClient;
end;

procedure TIWForm2.rgnintroCreate(Sender: TObject);
begin
rgnintro.align := AlClient;
end;

procedure TIWForm2.rgnloginCreate(Sender: TObject);
begin
 rgnlogin.Align := AlClient;
end;

procedure TIWForm2.rgnnovoproblemaCreate(Sender: TObject);
begin
rgnnovoproblema.Align := alclient;
end;

procedure TIWForm2.rgnOpcaoInputCreate(Sender: TObject);
begin
rgnopcaoinput.Align := AlClient;
end;

procedure TIWForm2.rgrpPeriodoClick(Sender: TObject);
begin
With TIWUserSession(WebApplication.Data) do
begin
if rgrpPeriodo.ItemIndex = 0 then
Periodo := 12;
if rgrpPeriodo.ItemIndex = 1 then
Periodo := 6;
if rgrpPeriodo.ItemIndex = 2 then
Periodo := 4;

if (lstbxperformanceobj.Items.Count > 2) and (btnsalvaretapa1.Enabled = False) then
btnavancaretapa1.Enabled := True;
end;
end;



procedure TIWForm2.edtestoqAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtestoq.Text[Length(edtestoq.Text)] <> '0') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '1') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '2') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '3') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '4') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '5') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '6') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '7') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '8') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '9') then
    edtestoq.Text := Copy(edtestoq.Text, 1, length(edtestoq.Text) - 1)
    else
    edtestoq.Text := edtestoq.Text;
end;

procedure TIWForm2.edtfuncAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtfunc.Text[Length(edtfunc.Text)] <> '0') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '1') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '2') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '3') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '4') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '5') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '6') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '7') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '8') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '9') then
    edtfunc.Text := Copy(edtfunc.Text, 1, length(edtfunc.Text) - 1)
    else
    edtfunc.Text := edtfunc.Text;
end;

procedure TIWForm2.edthorasextrasAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edthorasextras.Text[Length(edthorasextras.Text)] <> '0') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '1') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '2') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '3') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '4') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '5') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '6') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '7') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '8') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '9') then
    edthorasextras.Text := Copy(edthorasextras.Text, 1, length(edthorasextras.Text) - 1)
    else
    edthorasextras.Text := edthorasextras.Text;
end;

procedure TIWForm2.edtinifuncAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtinifunc.Text[Length(edtinifunc.Text)] <> '0') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '1') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '2') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '3') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '4') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '5') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '6') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '7') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '8') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '9') then
    edtinifunc.Text := Copy(edtinifunc.Text, 1, length(edtinifunc.Text) - 1)
    else
    edtinifunc.Text := edtinifunc.Text;
end;



procedure TIWForm2.edtnewactorAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
edtnewactor.Clear;
end;

initialization
  TIWForm2.SetAsMainForm;

end.
