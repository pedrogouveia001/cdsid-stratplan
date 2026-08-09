unit untResultados;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompLabel,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, TeeGDIPlus, IWCompButton, TeEngine, ExtCtrls,
  TeeProcs, Chart, IWExtCtrls, IWWebGrid, IWAdvWebGrid, BubbleCh, TeePolar,
  IWCompListbox, IWCompRadioButton, IWTMSCheckList, series, IWGrids, TeeLisB,
  TeeSeriesTextEd, Buttons, Math, pngimage, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Windows, Messages, Variants, Graphics, Dialogs, ComCtrls, Menus, StdCtrls,
  jpeg, ComObj, Grids, IWCompText, IWHTMLControls, ErrorBar,
  IWAdvWebGridExcel, IWProgressIndicator, IWDBGrids,
  IWCompProgressBar, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWCompTabControl, IWTMSPopup, IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, IWHelpTip, IWCompCheckbox, GanttCh;
type
  TfrmResultados = class(TIWAppForm)
    rgnplanos: TIWRegion;
    IWImage6: TIWImage;
    IWLabel16: TIWLabel;
    IWLink7: TIWLink;
    IWLink9: TIWLink;
    IWLabel25: TIWLabel;
    IWRegion2: TIWRegion;
    cmbbxPlanos: TIWComboBox;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    IWRegion1: TIWRegion;
    grddesempenho: TIWGrid;
    btnsmarter: TIWButton;
    pmnExport: TTIWPopupMenuButton;
    pMnExporting: TPopupMenu;
    ConsequenceMatrix: TMenuItem;
    DetailedPlans: TMenuItem;
    grdmatrizconseq: TTIWAdvWebGrid;
    grdplans: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    TIWAdvWebGridExcelIO2: TTIWAdvWebGridExcelIO;
    rgnsmarter: TIWRegion;
    IWImage5: TIWImage;
    IWLink1: TIWLink;
    IWLink2: TIWLink;
    IWLabel4: TIWLabel;
    IWRegion4: TIWRegion;
    rgnprocroc: TIWRegion;
    IWRegion16: TIWRegion;
    IWLabel3: TIWLabel;
    rgnROC: TIWRegion;
    IWHRule1: TIWHRule;
    IWLabel26: TIWLabel;
    btngoroc: TIWButton;
    IWButton3: TIWButton;
    IWRegion11: TIWRegion;
    IWLabel51: TIWLabel;
    rgnpesosROC: TIWRegion;
    grdpesosROC: TIWGrid;
    IWRegion29: TIWRegion;
    lblNSOE: TIWLabel;
    txtS2OE: TIWText;
    IWRegion30: TIWRegion;
    lstbxrankROC: TIWListbox;
    lstbxcriterioROC: TIWListbox;
    lblCOSCOOE: TIWLabel;
    Btnrestart: TIWButton;
    btnchooseROC: TIWButton;
    IWRegion17: TIWRegion;
    imgGrafOverallE: TIWImage;
    btngerargraficoavhol: TIWButton;
    lblS1OE: TIWLabel;
    IWRegion31: TIWRegion;
    lblBiOE: TIWLabel;
    lblWiOE: TIWLabel;
    crtOverallE: TChart;
    Series4: TBarSeries;
    Series1: TBarSeries;
    rgnresultados: TIWRegion;
    IWImage9: TIWImage;
    lblmelhorplano: TIWLabel;
    IWLink3: TIWLink;
    IWLink4: TIWLink;
    IWLabel6: TIWLabel;
    TIWPopupMenuButton1: TTIWPopupMenuButton;
    IWRegion5: TIWRegion;
    tbctrlMelhorPlano: TIWTabControl;
    PageDetailedPlan: TIWTabPage;
    PagePerformancePlan: TIWTabPage;
    imgresultado: TIWImage;
    chrtResultado: TChart;
    BarSeries1: TBarSeries;
    BarSeries2: TBarSeries;
    btnPreenchimentoAS: TIWButton;
    IWRgnAS: TIWRegion;
    lblWSAL: TIWLabel;
    lblPISAL: TIWLabel;
    IWImage50: TIWImage;
    IWImage55: TIWImage;
    IWRegion3: TIWRegion;
    txtsesanw: TIWText;
    IWRegion47: TIWRegion;
    IWRgnanalisedesensibilidade: TIWRegion;
    imgsensgraph: TIWImage;
    lblDOSCSA: TIWLabel;
    btnexpRCSA: TIWButton;
    lblAIExcCSA: TIWLabel;
    lblVPVCSA: TIWLabel;
    lblCSA: TIWLabel;
    btnback: TIWButton;
    btnmximc: TIWButton;
    AdvWgrdASR: TTIWAdvWebGrid;
    IWRegion7: TIWRegion;
    txtAOSCSA: TIWText;
    TxtAnOSCSA: TIWText;
    IWImage10: TIWImage;
    IWImage12: TIWImage;
    lblLCSA: TIWLabel;
    chrtsensitivity: TChart;
    Series5: TBarSeries;
    BarSeries3: TBarSeries;
    IWRegion6: TIWRegion;
    grddeviation: TIWGrid;
    IWRegion8: TIWRegion;
    grdaltinout: TIWGrid;
    rgnvariationconsequences: TIWRegion;
    grdcrivarMConseq: TIWGrid;
    IWRegion40: TIWRegion;
    IWImage11: TIWImage;
    IWImage13: TIWImage;
    rgnsav: TIWRegion;
    lblSAV: TIWLabel;
    lblSOCSAV: TIWLabel;
    btnsaveas: TIWButton;
    btnbackAS: TIWButton;
    btnstartas: TIWButton;
    lnkBackAS: TIWLink;
    IWRegion46: TIWRegion;
    rgnconsequencias: TIWRegion;
    grdconsequencias: TTIWAdvWebGrid;
    IWImage14: TIWImage;
    IWLabel5: TIWLabel;
    rgnpesos: TIWRegion;
    grdpesos: TTIWAdvWebGrid;
    IWTimer1: TIWTimer;
    IWLabel7: TIWLabel;
    rgnvariationweight: TIWRegion;
    grdcrivarWeight: TIWGrid;
    rgnimagemax: TIWRegion;
    imagemax: TIWImage;
    btnclosemax: TIWButton;
    grdexportacaoAS: TTIWAdvWebGrid;
    IWRegion13: TIWRegion;
    IWImage26: TIWImage;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWRegion9: TIWRegion;
    IWImage7: TIWImage;
    IWImage8: TIWImage;
    IWRegion14: TIWRegion;
    IWImage15: TIWImage;
    IWRegion12: TIWRegion;
    IWImage16: TIWImage;
    IWImage4: TIWImage;
    IWImage17: TIWImage;
    IWRegion10: TIWRegion;
    IWImage18: TIWImage;
    rgnplanosporperiodo: TIWRegion;
    grdplanos: TIWGrid;
    IWRegion15: TIWRegion;
    grdresultado: TTIWAdvWebGrid;
    procedure rgnplanosCreate(Sender: TObject);
    procedure cmbbxPlanosChange(Sender: TObject);
    procedure ConsequenceMatrixClick(Sender: TObject);
    procedure DetailedPlansClick(Sender: TObject);
    procedure btnsmarterClick(Sender: TObject);
    procedure ClassificarPlanos;
    procedure gerar_grafico_overall;
    procedure Intra_Razao;
    Function casas_decimais(Num: Double): String;
    procedure rgnsmarterCreate(Sender: TObject);
    procedure btnchooseROCClick(Sender: TObject);
    procedure CalculaROC;
    procedure BtnrestartClick(Sender: TObject);
    procedure lstbxcriterioROCClick(Sender: TObject);
    procedure IWButton3Click(Sender: TObject);
    procedure btngorocClick(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure rgnresultadosCreate(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure IWLink9Click(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure grdconsequenciasCheckClick(Sender: TObject; RowIndex,
      ColumnIndex: Integer);
    procedure CriarElementosAS;
    procedure DestruirElementosAS;
    procedure btnPreenchimentoASClick(Sender: TObject);
    procedure rgnsavCreate(Sender: TObject);
    procedure grdpesosCheckClick(Sender: TObject; RowIndex,
      ColumnIndex: Integer);
    procedure lnkBackASClick(Sender: TObject);
    procedure btnsaveasClick(Sender: TObject);
    procedure btnbackASClick(Sender: TObject);
    procedure btnstartasClick(Sender: TObject);
    procedure IWTimer1Timer(Sender: TObject);
    procedure SensitivityA;
    function UniformPDF(min, max: real): real;
    function TrocaVP(Valor: string): String;
    function TrocaPV(Valor: string): String;
    procedure IWRgnanalisedesensibilidadeCreate(Sender: TObject);
    procedure btnbackClick(Sender: TObject);
    procedure btnclosemaxClick(Sender: TObject);
    procedure rgnimagemaxCreate(Sender: TObject);
    procedure btnmximcClick(Sender: TObject);
    procedure IWRgnASCreate(Sender: TObject);
    procedure btnexpRCSAClick(Sender: TObject);
  public
  end;

implementation

uses Unit2, UserSessionUnit, ServerController;

{$R *.dfm}

Function TfrmResultados.UniformPDF(min, max: real): real;
var
  aux: real;
begin
  aux := Random;
  Result := (aux * (max - min)) + min;
end;

function TfrmResultados.TrocaVP(Valor: string): String;
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

function TfrmResultados.TrocaPV(Valor: string): String;
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

procedure TfrmResultados.ClassificarPlanos;
var
  i, j: integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  Ncrit := length(ObjSelecionados);
  Nalt := length(populacao);
  SetLength(TIWUserSession(WebApplication.Data).maior, TIWUserSession
      (WebApplication.Data).Ncrit);
  SetLength(TIWUserSession(WebApplication.Data).menor, TIWUserSession
      (WebApplication.Data).Ncrit);
  for j := 1 to TIWUserSession(WebApplication.Data).Ncrit do
  begin
    for i := 1 to TIWUserSession(WebApplication.Data).Nalt do
    begin
      if i = 1 then
      begin
        TIWUserSession(WebApplication.Data).maior[j - 1] := TIWUserSession
          (WebApplication.Data).PerformancePopulacao[i - 1, j - 1];
        TIWUserSession(WebApplication.Data).menor[j - 1] := TIWUserSession
          (WebApplication.Data).PerformancePopulacao[i - 1, j - 1];
      end
      else
      begin
        if TIWUserSession(WebApplication.Data).PerformancePopulacao[i - 1, j - 1]
          > TIWUserSession(WebApplication.Data).maior[j - 1] then
          TIWUserSession(WebApplication.Data).maior[j - 1] := TIWUserSession
            (WebApplication.Data).PerformancePopulacao[i - 1, j - 1];
        if TIWUserSession(WebApplication.Data).PerformancePopulacao[i - 1, j - 1]
          < TIWUserSession(WebApplication.Data).menor[j - 1] then
          TIWUserSession(WebApplication.Data).menor[j - 1] := TIWUserSession
            (WebApplication.Data).PerformancePopulacao[i - 1, j - 1];
      end;
    end;
  end;
  end;

end;

procedure TfrmResultados.btnbackASClick(Sender: TObject);
var
i : integer;
begin
 with TIWUserSession(WebApplication.Data) do
begin
  Finalize(pmaxMConseq);
  finalize(pminMConseq);
  finalize(criselMConseq);
  for i := 0 to grdconsequencias.RowCount - 1 do
  begin
    if grdconsequencias.RowSelect[i] = True then
    begin
      edtmaxMConseq[i].enabled := True;
      edtminMConseq[i].enabled := True;
    end;
  end;
  Finalize(pmaxWeight);
  finalize(pminWeight);
  finalize(criselWeight);
  for i := 0 to grdpesos.RowCount - 1 do
  begin
    if grdconsequencias.RowSelect[i] = True then
    begin
      edtmaxWeight[i].enabled := True;
      edtminWeight[i].enabled := True;
    end;
  end;

end;
btnsaveAS.Enabled := true;
btnbackAS.Enabled := false;
btnstartas.Enabled := False;
grdconsequencias.Enabled := true;
grdpesos.Enabled := True;
end;

procedure TfrmResultados.btnbackClick(Sender: TObject);
begin
IWRgnanalisedesensibilidade.Visible := False;
rgnresultados.Visible := True;
grdcrivarMConseq.Clear;
grdcrivarWeight.Clear;
grdaltinout.Clear;
grddeviation.Clear;
end;

procedure TfrmResultados.btnchooseROCClick(Sender: TObject);
var
  i: integer;
begin
  TIWUserSession(WebApplication.Data).vetorGrafico[lstbxcriterioROC.ItemIndex]
    := 2;
  gerar_grafico_overall;
  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(indexroc, Ncrit); // ajustar
    if (lstbxcriterioROC.ItemIndex <> -1) and
      (lstbxcriterioROC.Items[lstbxcriterioROC.ItemIndex] <> '') then
    begin
      lstbxrankROC.Items.Add(lstbxcriterioROC.Items[lstbxcriterioROC.ItemIndex]
        );
      indexroc[lstbxrankROC.Items.Count - 1] := lstbxcriterioROC.ItemIndex;
      lstbxcriterioROC.Items[lstbxcriterioROC.ItemIndex] := '';
      if lstbxrankROC.Items.Count = Ncrit - 1 then
      begin
        for i := 0 to Ncrit - 1 do
          if lstbxcriterioROC.Items[i] <> '' then
          begin
            lstbxrankROC.Items.Add(lstbxcriterioROC.Items[i]);
            indexroc[lstbxrankROC.Items.Count - 1] := i;
            lstbxcriterioROC.Items[i] := '';
            TIWUserSession(WebApplication.Data).vetorGrafico[i] := 2;
          end;
        btnchooseROC.Enabled := false;
        btngoroc.Enabled := true;
        gerar_grafico_overall;
        // WebApplication.ShowMessage('ROC procedure completed', smAlert);
        CalculaROC;
        rgnROC.Visible := true;
      end;
    end;
  end;

end;

procedure TfrmResultados.btnclosemaxClick(Sender: TObject);
begin
rgnimagemax.Hide;
end;

procedure TfrmResultados.btnexpRCSAClick(Sender: TObject);
var
i, j : integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    i := grddeviation.ColumnCount + grdaltinout.ColumnCount + grdcrivarMConseq.ColumnCount;
    for j := 0 to i do
    grdexportacaoAS.Columns.Add;

    grdexportacaoAS.RowCount := grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount + grddeviation.RowCount + grdaltinout.RowCount + 20;
    grdexportacaoAS.TotalRows := grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount + grddeviation.RowCount + grdaltinout.RowCount + 20;


    grdexportacaoAS.Cells[0,0] := 'Sensitivity Analysis Results';
    grdexportacaoAS.Cells[0,2] := 'Consequences Variation:';
    for i := 0 to grdcrivarMConseq.RowCount  - 1 do
    for j := 0 to grdcrivarMconseq.ColumnCount - 1 do
    begin
      grdexportacaoAS.Cells[j,i+3] := grdcrivarMconseq.Cell[i,j].Text;
    end;

    grdexportacaoAS.Cells[0, grdcrivarMConseq.RowCount +4] := 'Weigths Variation:';
    for i := 0 to grdcrivarWeight.RowCount  - 1 do
    for j := 0 to grdcrivarWeight.ColumnCount - 1 do
    begin
      grdexportacaoAS.Cells[j,i + grdcrivarMConseq.RowCount + 5] := grdcrivarWeight.Cell[i,j].Text;
    end;

    grdexportacaoAS.Cells[0, grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount +6] := 'Deviation from the original solution:';
    for i := 0 to grddeviation.RowCount  - 1 do
    for j := 0 to grddeviation.ColumnCount - 1 do
    begin
      grdexportacaoAS.Cells[j,i + grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount +7] := grddeviation.Cell[i,j].Text;
    end;

    grdexportacaoAS.Cells[0, grddeviation.RowCount+grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount +8] := 'Alternatives Included as a solution throughout the interactions of the Sensitivity Analysis:';
    for i := 0 to grdaltinout.RowCount  - 1 do
    for j := 0 to grdaltinout.ColumnCount - 1 do
    begin
      grdexportacaoAS.Cells[j,i + grddeviation.RowCount + grdcrivarMConseq.RowCount + grdcrivarWeight.RowCount +9] := grdaltinout.Cell[i,j].Text;
    end;

   TIWAdvWebGridExcelIO1.AdvWebGrid:= grdexportacaoAS;
   TIWAdvWebGridExcelIO1.XLSExport('Sensitivity Analysis Results - ' + problema + '.xls', WebApplication);
  end;
end;

procedure TfrmResultados.CalculaROC;
var
  i, j: integer;
  peso, soma: real;
begin
  for i := 0 to lstbxrankROC.Items.Count - 1 do
  begin
    soma := 0;
    for j := i + 1 to TIWUserSession(WebApplication.Data).Ncrit do
    begin
      soma := soma + 1 / j;
    end;
    peso := (1 / TIWUserSession(WebApplication.Data).Ncrit) * soma;
    grdpesosROC.Cell[2, TIWUserSession(WebApplication.Data).indexroc[i] + 1]
      .Text := FloatToStr(RoundTo(peso, -3));
  end;
end;

procedure TfrmResultados.btngorocClick(Sender: TObject);
var
  i, j: integer;
  Valor, temp: real;
  tempindex: integer;
  resultado : array of real;
  performanceMedia, performanceMediaNorm : array of real;
  ResultadoMedio : real;
begin
with TIWUserSession(WebApplication.Data) do
begin
 SetLength(TIWUserSession(WebApplication.Data).peso, TIWUserSession
      (WebApplication.Data).Ncrit);
  for i := 1 to TIWUserSession(WebApplication.Data).Ncrit do
  begin
    TIWUserSession(WebApplication.Data).peso[i - 1] := StrToFloat
      (grdpesosROC.Cell[2, i].Text);
  end;
  SetLength(PerformanceNormalizada, Nalt, Ncrit);
  for i := 1 to Nalt do
    for j := 1 to Ncrit do
  begin
  if maior[j-1] = menor[j-1]  then
  TIWUserSession(WebApplication.Data).PerformanceNormalizada[i - 1, j - 1] := 0
  else
  begin
    if objSelecionados[j-1] <> 'Cost' then
        TIWUserSession(WebApplication.Data).PerformanceNormalizada[i - 1, j - 1] :=(PerformancePopulacao[i - 1, j - 1] - TIWUserSession(WebApplication.Data).menor[j - 1]) /
          (maior[j - 1] - menor[j - 1])
      else if objSelecionados[j-1] = 'Cost' then
        PerformanceNormalizada[i - 1, j - 1] :=(PerformancePopulacao[i - 1, j - 1] - maior[j - 1]) /
          (menor[j - 1] - maior[j - 1]);
  end;
  end;
  SetLength(Resultado, Nalt);
  valor := 0;
  for i := 0 to Nalt - 1 do
  begin
    Resultado[i] := 0;
    for J := 0 to Ncrit - 1 do
    Resultado[i] := Resultado[i] + PerformanceNormalizada[i,j]*Peso[j];
    if Resultado[i] > valor then
    begin
      valor := resultado[i];
      tempindex := i;
    end;
  end;
melhorplano := tempindex;


    SetLength(performancemedia, Ncrit);
    SetLength(performancemedianorm, Ncrit);
    for j := 0 to Ncrit - 1 do
    begin
      for i := 0 to Nalt - 1 do
      begin
        PerformanceMedia[j] := PerformancePopulacao[i, j] + PerformanceMedia[j]
      end;
      PerformanceMedia[j] := PerformanceMedia[j]/Nalt;
      if maior[j] = menor[j] then
      PerformanceMediaNorm[j] := 0
      else
      begin
      if objSelecionados[j] <> 'Cost' then
      PerformanceMediaNorm[j] := (PerformanceMedia[j] - menor[j]) / (maior[j] - menor[j])
      else if objSelecionados[j] = 'Cost' then
      PerformanceMediaNorm[j] := (PerformanceMedia[j] - maior[j]) / (menor[j] - maior[j]);
      end;
      ResultadoMedio := ResultadoMedio + PerformanceMediaNorm[j]*Peso[j];
    end;





for i := 0 to Length(populacao[0]) - 1 do
begin
grdresultado.Cells[0,i] := inttostr(i+1);
for j := 0 to Length(populacao[0,0]) - 1 do
grdresultado.Cells[j+1,i] := floattostr(populacao[melhorplano,i,j]);
end;
 // torna o componente IWImage visivel
  chrtResultado.BottomAxis.Visible := true; // torna o eixo x visível
  chrtResultado.SeriesList.ClearValues;
  // limpar as séries se ja houver algum dado
  chrtResultado.leftAxis.Items.Clear; // limpando os valores dos eixos para não se sobrepor

  for i := 0 to Ncrit - 1 do
  begin
    chrtResultado.Series[0].AddXY(i, PerformanceNormalizada[melhorplano,i], casas_decimais(PerformancePopulacao[melhorplano,i]));
    chrtResultado.Series[1].AddXY(i, PerformanceMediaNorm[i], casas_decimais(PerformanceMedia[i]));
    chrtResultado.BottomAxis.Items.Add(i, ObjSelecionados[i]);
  end;
  chrtResultado.Series[0].AddXY(Ncrit,  valor, casas_decimais(valor), clWebDarkBlue);
  chrtResultado.Series[1].AddXY(Ncrit,  resultadoMedio, casas_decimais(resultadoMedio), clWebDarkRed);
  chrtResultado.BottomAxis.Items.add(Ncrit, 'Overall Performance');
  chrtResultado.Title.Caption := 'Plan '  + inttostr(melhorplano+1) + 'vs Average Performance';
  chrtresultado.Legend.Visible := True;
  chrtResultado.Series[0].Legend.text := 'Plan '  + inttostr(melhorplano+1);
  chrtResultado.Series[1].Legend.text := 'Average Performance';
  imgresultado.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtResultado.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgresultado.Width,
          imgresultado.Height)));
  rgnsmarter.Visible := False;
  rgnresultados.Visible := True;
  lblmelhorplano.Caption := 'Plan '  + inttostr(melhorplano+1) + ' was selected as the best plan:';
 WebApplication.ShowMessage('The problem has been solved successfully!');
end;

end;


procedure TfrmResultados.btnmximcClick(Sender: TObject);
begin
rgnimagemax.Show;
end;

procedure TfrmResultados.btnPreenchimentoASClick(Sender: TObject);
begin
rgnresultados.Visible := False;
rgnsav.Visible := True;
CriarElementosAS;
end;

procedure TfrmResultados.BtnrestartClick(Sender: TObject);
begin
  lstbxcriterioROC.Clear;
  lstbxrankROC.Clear;
  rgnROC.Visible := false;
  btnsmarterClick(Sender);
end;

procedure TfrmResultados.btnsaveasClick(Sender: TObject);
var
i, a, b  : integer;
verifica : boolean;
verifica2 : boolean;
begin
 with TIWUserSession(WebApplication.Data) do
  begin
  Finalize(pmaxMConseq); // Mudança - Pedro - 01.12
  Finalize(pminMConseq); // Mudança - Pedro - 01.12
  finalize(criselMConseq); // Mudança - Pedro - 01.12
  setlength(pmaxMConseq, Ncrit);
  setlength(pminMConseq, Ncrit);
  setlength(criselMConseq , Ncrit);

  Finalize(pmaxWeight); // Mudança - Pedro - 01.12
  Finalize(pminWeight); // Mudança - Pedro - 01.12
  finalize(criselWeight); // Mudança - Pedro - 01.12
  setlength(pmaxWeight, Ncrit);
  setlength(pminWeight, Ncrit);
  setlength(criselWeight , Ncrit);

  //ncv := 1;
  verifica2 := False;
  //PREENCHIMENTO DA VARIAÇÃO DAS CONSEQUENCIAS
  for i := 0 to grdconsequencias.RowCount - 1 do
  begin
    criselMConseq[i] := False;
    if grdconsequencias.RowSelect[i] = True then
    begin
      verifica2 := True;
      criselMConseq[i] := true;
        verifica := false;
      for a := 1 to length(edtmaxMConseq[i].Text) do
      for b := 0 to 9 do
      begin
        if edtMaxMConseq[i].Text[a] = inttostr(b) then
        begin
          verifica := true;
          break;
        end;
      end;
      if verifica = false then
      begin
        WebApplication.ShowMessage('Please fill in the parameters correctly!');
        exit;
      end;
      verifica := false;
      for a := 1 to length(edtminMConseq[i].Text) do
      for b := 0 to 9 do
      begin
        if edtMinMConseq[i].Text[a] = inttostr(b) then
        begin
          verifica := true;
          break;
        end;
      end;
      if verifica = false then
      begin
        WebApplication.ShowMessage('Please fill in the parameters correctly!');
        exit;
      end;
      edtmaxMConseq[i].enabled := False;
      edtminMConseq[i].enabled := False;
      edtmaxMConseq[i].Text := TrocaVP(edtmaxMConseq[i].Text);
      edtminMConseq[i].Text := TrocaVP(edtminMConseq[i].Text);
      pmaxMConseq[i] := strtofloat(edtmaxMConseq[i].Text);
      pminMConseq[i] := strtofloat(edtminMConseq[i].Text);
    end;
  end;

  // PREENCHIMENTO DA VARIAÇÃO DOS PESOS
  for i := 0 to grdpesos.RowCount - 1 do
  begin
    criselWeight[i] := False;
    if grdpesos.RowSelect[i] = True then
    begin
      verifica2 := True;
      criselWeight[i] := true;
        verifica := false;
      for a := 1 to length(edtmaxWeight[i].Text) do
      for b := 0 to 9 do
      begin
        if edtMaxWeight[i].Text[a] = inttostr(b) then
        begin
          verifica := true;
          break;
        end;
      end;
      if verifica = false then
      begin
        WebApplication.ShowMessage('Please fill in the parameters correctly!');
        exit;
      end;
      verifica := false;
      for a := 1 to length(edtminWeight[i].Text) do
      for b := 0 to 9 do
      begin
        if edtMinWeight[i].Text[a] = inttostr(b) then
        begin
          verifica := true;
          break;
        end;
      end;
      if verifica = false then
      begin
        WebApplication.ShowMessage('Please fill in the parameters correctly!');
        exit;
      end;
      edtmaxWeight[i].enabled := False;
      edtminWeight[i].enabled := False;
      edtmaxWeight[i].Text := TrocaVP(edtmaxWeight[i].Text);
      edtminWeight[i].Text := TrocaVP(edtminWeight[i].Text);
      pmaxWeight[i] := strtofloat(edtmaxWeight[i].Text);
      pminWeight[i] := strtofloat(edtminWeight[i].Text);
    end;
  end;
end; // With
if verifica2 = false then
begin
WebApplication.ShowMessage('Please select a criterion to continue!');
exit;
end;
btnsaveAS.Enabled := False;
btnbackAS.Enabled := true;
btnstartAS.enabled := True;
grdconsequencias.Enabled := False;
end;

procedure TfrmResultados.btnsmarterClick(Sender: TObject);
var
  i: integer;
begin
  ClassificarPlanos;
  lstbxcriterioROC.Items.Clear;
  lstbxrankROC.Items.Clear;
  grdpesosROC.Clear;
  btnchooseROC.Enabled := true;
  btngoroc.Enabled := false;
  With TIWUserSession(WebApplication.Data) do
  begin
  with grdpesosROC do
  begin
    RowCount := 3;
    ColumnCount := Length(ObjSelecionados) + 1;
    Cell[0, 0].Text := 'Criterion´s Code';
    Cell[1, 0].Text := 'Criterion´s Name';
    Cell[2, 0].Text := 'ROC weight';
  end;

  for i := 1 to length(ObjSelecionados) do
  begin
    grdpesosROC.Cell[0, i].Text := 'Crit ' + IntToStr(i);
    grdpesosROC.Cell[1, i].Text := ObjSelecionados[i - 1];
    grdpesosROC.Cell[2, i].Text := '';
    lstbxcriterioROC.Items.Add
      (grdpesosROC.Cell[0, i].Text + ': ' + grdpesosROC.Cell[1, i].Text);
  end;

  SetLength(vetorGrafico, Length(ObjSelecionados));
  for i := 1 to Length(ObjSelecionados) do
    TIWUserSession(WebApplication.Data).vetorGrafico[i - 1] := 0;
  gerar_grafico_overall;
end;
rgnSmarter.Visible := True;
rgnplanos.Visible := False;
end;

procedure TfrmResultados.btnstartasClick(Sender: TObject);
begin
IWRgnAS.Visible := True;
rgnsav.Visible := false;
DestruirElementosAS;
btnSaveAS.Enabled := true;
btnbackas.Enabled := False;
btnstartas.Enabled := False;
grdconsequencias.Enabled := true;
grdpesos.Enabled := True;
TIWUserSession(WebApplication.Data).Start := True;
TIWUserSession(WebApplication.Data).Finish := false;
IWTimer1.Enabled := True;
end;

procedure TfrmResultados.gerar_grafico_overall;
var
  i: integer;
  bt1, wt1, Bt: string;
begin
  // ** Caso passe a ser bilingue, essas variáveis seriam mudadas **//
  bt1 := 'B';
  wt1 := 'W';
  Bt := 'Best';
  Intra_Razao;
  imgGrafOverallE.Visible := true;
  // torna o componente IWImage visivel
  crtOverallE.BottomAxis.Visible := true; // torna o eixo x visível
  crtOverallE.SeriesList.ClearValues;
  // limpar as séries se ja houver algum dado
  crtOverallE.leftAxis.Items.Clear; // limpando os valores dos eixos para não se sobrepor

  with TIWUserSession(WebApplication.Data) do
  begin
    crtOverallE.leftAxis.Items.Add(100, Bt); // adicionando o nome best no eixo y do melhor cirtério
    for i := 0 to Ncrit - 1 do
    begin
      crtOverallE.BottomAxis.Items.Add(i, 'Crit ' + IntToStr(i + 1));
      if vetorGrafico[i] = 1 then
      // ** O critério está selecionado no listbox  **//
      begin
        if ObjSelecionados[i] <> 'Cost' then
        crtOverallE.Series[1].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .BConseqOrd[i]), clwebYELLOW)
        else
         crtOverallE.Series[1].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]), clwebYELLOW)
      end;

      if vetorGrafico[i] = 2 then
      // ** Critério ja ordenado **//
      begin
        if ObjSelecionados[I] <> 'Cost' then
        crtOverallE.Series[1].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .BConseqOrd[i]), clWebGREEN)
        else
         crtOverallE.Series[1].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]), clWebGREEN)
      end;
      if vetorGrafico[i] = 0 then
      begin
        // ** O critério ainda não foi nem selecionado, nem ordenado **//
        if i <> 0 then
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .BConseqOrd[i]))
        else
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]));

        if i = 0 then
        crtOverallE.Series[1].AddXY(i, TIWUserSession(WebApplication.Data)
            .valor_razao[i] * 100, wt1 + IntToStr(i + 1) + ':' + casas_decimais
            (TIWUserSession(WebApplication.Data).BConseqOrd[i]), clWebRED)
        else
        crtOverallE.Series[1].AddXY(i, TIWUserSession(WebApplication.Data)
            .valor_razao[i] * 90, wt1 + IntToStr(i + 1) + ':' + casas_decimais
            (TIWUserSession(WebApplication.Data).WConseqOrd[i]), clWebRED)
      end;
    end;

    imgGrafOverallE.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (crtOverallE.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgGrafOverallE.Width,
          imgGrafOverallE.Height)));
  end;
end;

Procedure TfrmResultados.DestruirElementosAS;
var
i, j : integer;
begin
//rgnsav.Align := AlClient;
btnsaveas.enabled := True;
btnbackAS.enabled := False;
btnstartas.Enabled := False;
with TIWUserSession(WebApplication.Data) do
begin
  for i := 0 to grdconsequencias.RowCount - 1 do
  begin
   lblmaxMConseq[i].Destroy;
   lblminMConseq[i].Destroy;

   if grdconsequencias.RowSelect[i] = True then
   grdconsequencias.RowSelect[i] := False;

   edtmaxMConseq[i].Destroy;
   edtminMConseq[i].Destroy;
   lblvariationminMConseq[i].Destroy;
   lblvariationmaxMConseq[i].Destroy;
  end;
  for i := 0 to grdpesos.RowCount - 1 do
  begin
   lblmaxWeight[i].Destroy;
   lblminWeight[i].Destroy;

   if grdpesos.RowSelect[i] = True then
   grdpesos.RowSelect[i] := False;

   edtmaxWeight[i].Destroy;
   edtminWeight[i].Destroy;
   lblvariationminWeight[i].Destroy;
   lblvariationmaxWeight[i].Destroy;
  end;
end;
end;

Procedure TfrmResultados.CriarElementosAS;
var
NumAlt, Numcrit : integer;
NomeAlt, Nomecrit : string;
i, j : integer;
begin
//rgnsav.Align := AlClient;
btnsaveas.enabled := True;
btnbackAS.enabled := False;
btnstartas.Enabled := False;
with TIWUserSession(WebApplication.Data) do
begin
  //CRIAÇÃO DOS ELEMENTOS DA VARIAÇÃO DAS CONSEQUENCIAS
  grdconsequencias.RowCount := Length(ObjSelecionados);
  grdconsequencias.Columns[0].ColumnType := ctCheckbox;



  grdconsequencias.Columns[0].enabled := true;
  SetLength(edtmaxMConseq, grdconsequencias.RowCount);
  SetLength(edtminMConseq, grdconsequencias.RowCount);
  SetLength(lblmaxMConseq, grdconsequencias.RowCount);
  SetLength(lblminMConseq, grdconsequencias.RowCount);
  SetLength(lblvariationmaxMConseq, grdconsequencias.RowCount);
  SetLength(lblvariationminMConseq, grdconsequencias.RowCount);

  for i := 0 to grdconsequencias.RowCount - 1 do
  begin
    grdconsequencias.Cells[1,i] := ObjSelecionados[i];
    lblminMConseq[i] := TiwLabel.Create(Self);
    lblminMConseq[i].Parent := rgnconsequencias;
    lblminMConseq[i].Left := 205 + 15;
    lblminMConseq[i].Top := 26 + i*23;
    lblminMConseq[i].Font.FontName := 'verdana';
    lblminMConseq[i].Caption := '-';
    //lblmin[i].Name := 'lblmin' + inttostr(i);
    lblminMConseq[i].Visible := False;
    lblminMConseq[i].ZIndex := 10000;

    lblmaxMConseq[i] := TiwLabel.Create(Self);
    lblmaxMConseq[i].Parent := rgnconsequencias;
    lblmaxMConseq[i].Left := 365 + 15;
    lblmaxMConseq[i].Top := 26 + i*23;
    lblmaxMConseq[i].Font.FontName := 'verdana';
    lblmaxMConseq[i].Caption := '+';
    //lblmax[i].Name := 'lblmax' + inttostr(i);
    lblmaxMConseq[i].Visible := False;
    lblmaxMConseq[i].ZIndex := 10000;


      edtminMConseq[i] := TiwEdit.Create(Self);
      edtminMConseq[i].Parent := rgnconsequencias;
      edtminMConseq[i].Left := 205 + 31;
      edtminMConseq[i].Top := 26 + i*23;
      edtminMConseq[i].Font.FontName := 'verdana';
      edtminMConseq[i].Height := 21;
      edtminMConseq[i].Width := 56;
      edtminMConseq[i].Caption := '';
      //edtminAS[i].Name := 'edtminAS' + inttostr(i);
      edtminMConseq[i].Visible := False;
      edtminMConseq[i].ZIndex := 10000;


      edtmaxMConseq[i] := TiwEdit.Create(Self);
      edtmaxMConseq[i].Parent := rgnconsequencias;
      edtmaxMConseq[i].Left := 365+31;
      edtmaxMConseq[i].Top := 26 + i*23;
      edtmaxMConseq[i].Font.FontName := 'verdana';
      edtmaxMConseq[i].Height := 21;
      edtmaxMConseq[i].Width := 56;
      edtmaxMConseq[i].Caption := '';
      //edtmaxAS[i].Name := 'edtmaxAS' + inttostr(i);
      edtmaxMConseq[i].Visible := False;
      edtmaxMConseq[i].ZIndex := 10000;

      lblvariationminMConseq[i] := TiwLabel.Create(Self);
      lblvariationminMConseq[i].Parent := rgnconsequencias;;
      lblvariationminMConseq[i].Left := 205 + 92;
      lblvariationminMConseq[i].Top := 26 + i*23;
      lblvariationminMConseq[i].Font.FontName := 'verdana';
      lblvariationminMConseq[i].Caption := '%';
      //lblvariationminMConseq[i].Name := 'lblvariationmin' + inttostr(i);
      lblvariationminMConseq[i].Visible := False;
      lblvariationminMConseq[i].ZIndex := 10000;

      lblvariationmaxMConseq[i] := TiwLabel.Create(Self);
      lblvariationmaxMConseq[i].Parent := rgnconsequencias;;
      lblvariationmaxMConseq[i].Left := 365+92;
      lblvariationmaxMConseq[i].Top := 26 + i*23;
      lblvariationmaxMConseq[i].Font.FontName := 'verdana';
      lblvariationmaxMConseq[i].Caption := '%';
      //lblvariationmax[i].Name := 'lblvariationmax' + inttostr(i);
      lblvariationmaxMConseq[i].Visible := False;
      lblvariationmaxMConseq[i].ZIndex := 10000;





end;
  // CRIAÇÃO DOS ELEMENTOS DA VARIAÇÃO DOS PESOS
  grdpesos.RowCount := Length(ObjSelecionados);
  grdpesos.Columns[0].ColumnType := ctCheckbox;



  grdpesos.Columns[0].enabled := true;
  SetLength(edtmaxWeight, grdpesos.RowCount);
  SetLength(edtminWeight, grdpesos.RowCount);
  SetLength(lblmaxWeight, grdpesos.RowCount);
  SetLength(lblminWeight, grdpesos.RowCount);
  SetLength(lblvariationmaxWeight, grdpesos.RowCount);
  SetLength(lblvariationminWeight, grdpesos.RowCount);

  for i := 0 to grdpesos.RowCount - 1 do
  begin
    grdpesos.Cells[1,i] := ObjSelecionados[i];
    lblminWeight[i] := TiwLabel.Create(Self);
    lblminWeight[i].Parent := rgnpesos;
    lblminWeight[i].Left := 205 + 15;
    lblminWeight[i].Top := 26 + i*23;
    lblminWeight[i].Font.FontName := 'verdana';
    lblminWeight[i].Caption := '-';
    //lblmin[i].Name := 'lblmin' + inttostr(i);
    lblminWeight[i].Visible := False;
    lblminWeight[i].ZIndex := 10000;

    lblmaxWeight[i] := TiwLabel.Create(Self);
    lblmaxWeight[i].Parent := rgnpesos;
    lblmaxWeight[i].Left := 365 + 15;
    lblmaxWeight[i].Top := 26 + i*23;
    lblmaxWeight[i].Font.FontName := 'verdana';
    lblmaxWeight[i].Caption := '+';
    //lblmax[i].Name := 'lblmax' + inttostr(i);
    lblmaxWeight[i].Visible := False;
    lblmaxWeight[i].ZIndex := 10000;


      edtminWeight[i] := TiwEdit.Create(Self);
      edtminWeight[i].Parent := rgnpesos;
      edtminWeight[i].Left := 205 + 31;
      edtminWeight[i].Top := 26 + i*23;
      edtminWeight[i].Font.FontName := 'verdana';
      edtminWeight[i].Height := 21;
      edtminWeight[i].Width := 56;
      edtminWeight[i].Caption := '';
      //edtminAS[i].Name := 'edtminAS' + inttostr(i);
      edtminWeight[i].Visible := False;
      edtminWeight[i].ZIndex := 10000;


      edtmaxWeight[i] := TiwEdit.Create(Self);
      edtmaxWeight[i].Parent := rgnpesos;
      edtmaxWeight[i].Left := 365+31;
      edtmaxWeight[i].Top := 26 + i*23;
      edtmaxWeight[i].Font.FontName := 'verdana';
      edtmaxWeight[i].Height := 21;
      edtmaxWeight[i].Width := 56;
      edtmaxWeight[i].Caption := '';
      //edtmaxAS[i].Name := 'edtmaxAS' + inttostr(i);
      edtmaxWeight[i].Visible := False;
      edtmaxWeight[i].ZIndex := 10000;

      lblvariationminWeight[i] := TiwLabel.Create(Self);
      lblvariationminWeight[i].Parent := rgnpesos;;
      lblvariationminWeight[i].Left := 205 + 92;
      lblvariationminWeight[i].Top := 26 + i*23;
      lblvariationminWeight[i].Font.FontName := 'verdana';
      lblvariationminWeight[i].Caption := '%';
      //lblvariationminWeight[i].Name := 'lblvariationmin' + inttostr(i);
      lblvariationminWeight[i].Visible := False;
      lblvariationminWeight[i].ZIndex := 10000;

      lblvariationmaxWeight[i] := TiwLabel.Create(Self);
      lblvariationmaxWeight[i].Parent := rgnpesos;
      lblvariationmaxWeight[i].Left := 365+92;
      lblvariationmaxWeight[i].Top := 26 + i*23;
      lblvariationmaxWeight[i].Font.FontName := 'verdana';
      lblvariationmaxWeight[i].Caption := '%';
      //lblvariationmax[i].Name := 'lblvariationmax' + inttostr(i);
      lblvariationmaxWeight[i].Visible := False;
      lblvariationmaxWeight[i].ZIndex := 10000;





end;
end;

end;

procedure TfrmResultados.grdconsequenciasCheckClick(Sender: TObject; RowIndex,
  ColumnIndex: Integer);
var
i : integer;
begin
with TIWUserSession(WebApplication.Data) do
begin
  for i := 0 to grdconsequencias.RowCount - 1 do
  begin
    if grdconsequencias.RowSelect[i] = True then
    begin
      edtmaxMConseq[i].Visible := true;
      edtminMConseq[i].Visible := True;
      lblmaxMConseq[i].Visible := True;
      lblminMConseq[i].Visible := True;
      lblvariationmaxMConseq[i].Visible := True;
      lblvariationminMConseq[i].Visible := True;
    end;
    if grdconsequencias.RowSelect[i] = False then
    begin
      edtmaxMConseq[i].Visible := False;
      edtminMConseq[i].Visible := False;
      lblmaxMConseq[i].Visible := False;
      lblminMConseq[i].Visible := False;
      lblvariationmaxMConseq[i].Visible := False;
      lblvariationminMConseq[i].Visible := False;
    end;
  end;
end;

end;

procedure TfrmResultados.grdpesosCheckClick(Sender: TObject; RowIndex,
  ColumnIndex: Integer);
var
i : integer;
begin
with TIWUserSession(WebApplication.Data) do
begin
  for i := 0 to grdpesos.RowCount - 1 do
  begin
    if grdpesos.RowSelect[i] = True then
    begin
      edtmaxWeight[i].Visible := true;
      edtminWeight[i].Visible := True;
      lblmaxWeight[i].Visible := True;
      lblminWeight[i].Visible := True;
      lblvariationmaxWeight[i].Visible := True;
      lblvariationminWeight[i].Visible := True;
    end;
    if grdpesos.RowSelect[i] = False then
    begin
      edtmaxWeight[i].Visible := False;
      edtminWeight[i].Visible := False;
      lblmaxWeight[i].Visible := False;
      lblminWeight[i].Visible := False;
      lblvariationmaxWeight[i].Visible := False;
      lblvariationminWeight[i].Visible := False;
    end;
  end;
end;

end;

Function TfrmResultados.casas_decimais(Num: Double): String;
var
  S: string;
  i, k: integer;
  inicia: Boolean;
begin
  if Num <= 10000 then
  begin
  // ** O número de casas decimais está definido para 2 **//
  inicia := false;
  k := 0;
  S := FloatToStr(Num);
  // ** Contando o número de casas decimais **//
  for i := 2 to Length(S) - 1 do
  begin
    if inicia = true then
      INC(k);
    if (S[i] = ',') or (S[i] = '.') then
      inicia := true;
  end;
  // ** Só precisa formatar se o número de casas decimais for maior que 2 **//
  if k >= 3 then
    Result := FormatFloat('0.0000', Num)
  else
    Result := S;
  end
  else
  begin
   k := 0;
   S := FloatToStr(Num);
   // ** Contando o número de casas decimais **//
  for i := 0 to Length(S) - 1 do
  begin
  if (S[i] = ',') or (S[i] = '.') then
  break;

  INC(k);
  end;
  Num := Num / Power(10,k);
  S := FormatFloat('0.0000', Num);
  Num := StrToFloat(S) * Power(10,k);
  Result := FloatToStr(num);
  end;
end;



procedure TfrmResultados.Intra_Razao;
var
  i: integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    // ** Setando o tamanho dos vetores **//
    SetLength(valor_razao, Ncrit);
    SetLength(BConseqOrd, Ncrit);
    SetLength(WConseqOrd, Ncrit);

    // ** Preenchendo vetores **//
    for i := 0 to Ncrit - 1 do
      // ** Critério de maximização **//
      if (ObjSelecionados[i] <> 'Custo') then
      begin
        BConseqOrd[i] := maior[i];
        WConseqOrd[i] := menor[i];
      end
      else
      begin
        // ** Critério de minimização **//
        BConseqOrd[i] := menor[i];
        WConseqOrd[i] := maior[i];
      end;

    for i := 0 to Ncrit - 1 do
    begin

      if WConseqOrd[i] > BConseqOrd[i] then
      /// condição de minimização ///
      begin
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] < 0) then
        // fórmula para caso negativo
        begin
          valor_razao[i] := (WConseqOrd[i] / BConseqOrd[i]);
        end;
        if (WConseqOrd[i] > 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso positivo
        begin
          valor_razao[i] := (BConseqOrd[i] / WConseqOrd[i]);
        end;
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := ((WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
              + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i])));
        end;
        if (BConseqOrd[i] < 0) and (WConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := ((WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
              + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i])));
        end;
      end; // do meu primeiro if

      if WConseqOrd[i] < BConseqOrd[i] then
      /// condição de maximização///
      begin
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] < 0) then
        // fórmula para caso negativo
        begin
          valor_razao[i] := BConseqOrd[i] / WConseqOrd[i];
        end;
        if (WConseqOrd[i] >= 0) and (BConseqOrd[i] >= 0) then
        // fórmula para caso positivo
        begin
          valor_razao[i] := WConseqOrd[i] / BConseqOrd[i];
        end;
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := (WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
            + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]));
        end;
        if (BConseqOrd[i] < 0) and (WConseqOrd[i] > 0) then
        begin
          valor_razao[i] := (WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
            + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]));
        end;
      end; // do 2º if
    end; // do for
  end; // do with
end;



procedure TfrmResultados.IWButton3Click(Sender: TObject);
begin
  lstbxcriterioROC.Clear;
  lstbxrankROC.Clear;
  rgnROC.Visible := false;
  btnsmarterClick(Sender);
end;
procedure TfrmResultados.IWLink1Click(Sender: TObject);
begin
rgnsmarter.Visible := False;
rgnplanos.Visible := True;
end;



procedure TfrmResultados.IWLink3Click(Sender: TObject);
begin
rgnplanos.Visible := True;
rgnresultados.Visible := False;
end;

procedure TfrmResultados.IWLink4Click(Sender: TObject);
begin
TIWUserSession(WebApplication.Data).reinicia := 1;
TiwForm2.Create(WebApplication).show;
end;

procedure TfrmResultados.IWLink7Click(Sender: TObject);
begin
TIWUserSession(WebApplication.Data).reinicia := 2;
TiwForm2.Create(WebApplication).show;
end;

procedure TfrmResultados.IWLink9Click(Sender: TObject);
begin
TIWUserSession(WebApplication.Data).reinicia := 1;
TiwForm2.Create(WebApplication).show;
end;

procedure TfrmResultados.IWRgnanalisedesensibilidadeCreate(Sender: TObject);
begin
IWRgnanalisedesensibilidade.Align := AlClient;
end;

procedure TfrmResultados.IWRgnASCreate(Sender: TObject);
begin
IWRgnAS.left := 5;
IWRgnAS.Top := 5;
end;

procedure TfrmResultados.SensitivityA;
var
  i, j, CasebyCase, auxCol, auxRow, tempcode, X, k: Integer;
  somapeso : real;
  Valor, temp: real;
  tempindex: integer;
  resultado : array of real;
  auxiliar, auxiliar2 : integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
  TIWUserSession(WebApplication.Data).NumCases := 1000;//StrToInt(edtNumCases.Text);
  // número de instâncias
  Finalize(PesoAS);
  Finalize(PerformancePopulacaoAS);
  Finalize(PerformanceNormalizadaAS);
  Finalize(RelatorioAS);
  //Finalize(MaiorAS);
  //Finalize(MenorAS);
  SetLength(PesoAS,Ncrit);
  SetLength(PerformancePopulacaoAS, Nalt, Ncrit);
  SetLength(PerformanceNormalizadaAS, Nalt, Ncrit);
  //SetLength(maioras, ncrit);
  //setlength(menoras, ncrit);
  setlength(RelatorioAS, Nalt);
  k := 1;
  while k <= TIWUserSession(WebApplication.Data).NumCases do
  begin
    // Variando os pesos segundo uma PDF

    for i := 1 to TIWUserSession(WebApplication.Data).Ncrit do
    begin
    if TIWUserSession(WebApplication.Data).criselWeight[i] = true then
    begin
        TIWUserSession(WebApplication.Data).pesoAS[i - 1] := UniformPDF
          ((TIWUserSession(WebApplication.Data).peso[i - 1] *
            (100 - TIWUserSession(WebApplication.Data).PminWeight[i-1]) / 100), TIWUserSession
            (WebApplication.Data).peso[i - 1] *
            (100 + TIWUserSession(WebApplication.Data).PmaxWeight[i-1]) / 100)

    end
    else
        TIWUserSession(WebApplication.Data).pesoAS[i - 1] := TIWUserSession
          (WebApplication.Data).peso[i - 1];
    end; // for i

    // normalizar os pesos
    somapeso := 0;
    for i := 1 to TIWUserSession(WebApplication.Data).Ncrit do
      somapeso := somapeso + TIWUserSession(WebApplication.Data).pesoAS[i - 1];
    for i := 1 to TIWUserSession(WebApplication.Data).Ncrit do
      TIWUserSession(WebApplication.Data).pesoAS[i - 1] := TIWUserSession
        (WebApplication.Data).pesoAS[i - 1] / somapeso;




   { // variando a MConseq segundo uma PDF
    for i := 0 to TIWUserSession(WebApplication.Data).Ncrit-1 do
    begin
      maioras[i] := -99999999999;
      menorAS[i] := 999999999999;
    end;}

    for j := 0 to TIWUserSession(WebApplication.Data).Nalt-1 do
    begin
      for i := 0 to TIWUserSession(WebApplication.Data).Ncrit-1 do
      begin
        if TIWUserSession(WebApplication.Data).criselMConseq[i] = true then
        begin
          TIWUserSession(WebApplication.Data).PerformancePopulacaoAS[j, i] :=
          UniformPDF(PerformancePopulacao[j, i] * (100 - PminMConseq[i]) / 100,
          PerformancePopulacao[j, i] * (100 + PmaxMConseq[i]) / 100);
          if  PerformancePopulacaoAS[j, i] > maior[i] then
          PerformancePopulacaoAS[j, i] := maior[i];
          if  PerformancePopulacaoAS[j, i] < menor[i] then
          PerformancePopulacaoAS[j, i] := menor[i];
        end
        else
        begin
          TIWUserSession(WebApplication.Data).PerformancePopulacaoAS[j, i] :=
          TIWUserSession(WebApplication.Data).PerformancePopulacao[j, i];
        end;
        {if PerformancePopulacaoAS[j,i] > maiorAS[i] then
        maiorAS[i] := PerformancePopulacaoAS[j,i];

        if PerformancePopulacaoAS[j,i] < menorAS[i] then
        menorAS[i] := PerformancePopulacaoAS[j,i];
        }
      end
    end; // for i


  for i := 0 to Nalt -1 do
    for j := 0 to Ncrit -1 do
  begin
  if maior[j] = menor[j] then
  TIWUserSession(WebApplication.Data).PerformanceNormalizadaAS[i, j] := 0
  else
  begin
    if objSelecionados[j] <> 'Cost' then
        TIWUserSession(WebApplication.Data).PerformanceNormalizadaAS[i, j] :=(PerformancePopulacaoAS[i, j] - TIWUserSession(WebApplication.Data).menor[j]) /
          (maior[j] - menor[j])
      else if objSelecionados[j] = 'Cost' then
        PerformanceNormalizadaAS[i, j] :=(PerformancePopulacaoAS[i, j] - maior[j]) /
          (menor[j] - maior[j]);
  end;
  end;

  SetLength(Resultado, Nalt);
  valor := 0;
  for i := 0 to Nalt - 1 do
  begin
    Resultado[i] := 0;
    for J := 0 to Ncrit - 1 do
    Resultado[i] := Resultado[i] + TIWUserSession(WebApplication.Data).PerformanceNormalizadaAS[i,j]*PesoAS[j];
    if Resultado[i] > valor then
    begin
      valor := resultado[i];
      tempindex := i;
    end;
  end;
  melhorplanoAS := tempindex;

  inc(TIWUserSession(WebApplication.Data).RelatorioAS[melhorplanoAS]);
  inc(k);
  end;
  Start := False;
  Finish := True;
  chrtsensitivity.series[0].Clear;
  chrtsensitivity.series[1].Clear;
  grdaltinout.ColumnCount := 2;
  grdaltinout.RowCount := 2;
  grddeviation.ColumnCount := 3;
  grddeviation.RowCount := 2;
  grdaltinout.Cell[0,0].Font.Style := [fsbold];
  grdaltinout.Cell[0,1].Font.Style := [fsbold];
   grddeviation.Cell[0, 0].Text := 'Alternative';
        grddeviation.Cell[0, 1].Text := '% Included';
        grddeviation.Cell[0, 2].Text := '% Excluded';
        grdaltinout.Cell[0, 0].Text :=
          'Alternative';
        grdaltinout.Cell[0, 1].Text := '% Ocurrence';

      for j := 0 to grddeviation.ColumnCount - 1 do
      begin
        grddeviation.cell[0,j].Alignment := taCenter;
        grddeviation.cell[0,j].wrap := True;
        grddeviation.Cell[0,j].Font.Style := [fsbold];
      end;

      for j := 0 to grdaltinout.ColumnCount - 1 do
      begin
        grdaltinout.cell[0,j].Alignment := taCenter;
        grdaltinout.cell[0,j].wrap := True;
        grdaltinout.Cell[0,j].Font.Style := [fsbold];
      end;

        grddeviation.cell[1,0].Text := 'Plan ' + inttostr(melhorplano+1);
        grddeviation.cell[1,1].Text := FormatFloat('0.00', 100*RelatorioAS[melhorplanoAS]/Numcases) + ' %';
        grddeviation.cell[1,2].Text := FormatFloat('0.00', 100-(100*RelatorioAS[melhorplanoAS]/Numcases)) + ' %';

        chrtsensitivity.series[0].AddXY(0, (100*RelatorioAS[melhorplanoAS]/Numcases),
                FormatFloat('0.00', 100*RelatorioAS[melhorplanoAS]/Numcases) + '%');

       chrtsensitivity.bottomaxis.Items.Add(0, 'Plan ' + inttostr(melhorplano+1));

        auxiliar := RelatorioAS[melhorplanoAS];


        grdaltinout.cell[1,0].Text := 'Plan ' + inttostr(melhorplano+1);
        grdaltinout.cell[1,1].Text := FormatFloat('0.00', 100*RelatorioAS[melhorplanoAS]/Numcases) + ' %';

        x := 1;
        j := 0;
        auxiliar2 := 0;
        for I := 0 to Nalt - 1 do
        begin
        if (i = melhorplano) then
        continue;

        if RelatorioAS[i] = 0  then
        continue;

        {if RelatorioAS[i] > auxiliar2 then
        begin
        auxiliar2 := RelatorioAS[i];
        j := i;
        end;}
        inc(x);
        grdaltinout.RowCount := x + 1;
        grdaltinout.Cell[x,0].Text := 'Plan ' + inttostr(i+1);
        grdaltinout.Cell[x,1].Text :=  FormatFloat('0.00', 100*RelatorioAS[i]/Numcases) + ' %';
        if RelatorioAS[i] > NumCases*0.05 then
        begin
        inc(j);
        chrtsensitivity.series[1].AddXY(j, (100*RelatorioAS[i]/Numcases),
                FormatFloat('0.00', 100*RelatorioAS[i]/Numcases) + '%');
        chrtsensitivity.bottomaxis.Items.Add(j, 'Plan ' + inttostr(i+1));
        //TIWUserSession(WebApplication.Data).auxiliar := auxiliar + relatorioAS[j];
       // TIWUserSession(WebApplication.Data).RelatorioAS[j] := 0;
        end;
        end;
        chrtsensitivity.LeftAxis.Increment := 0.2; // ?

      chrtsensitivity.BottomAxis.LabelsAlternate;
      chrtsensitivity.BottomAxis.Items.Automatic := false;
      chrtsensitivity.LeftAxis.Maximum := 100;
      chrtsensitivity.LeftAxis.MaximumOffset := 16;
      chrtsensitivity.LeftAxis.Minimum := 0;
      chrtsensitivity.LeftAxis.Increment := 20;
      chrtsensitivity.Visible := true;
      imgsensgraph.Visible := True;
      imgsensgraph.Picture.Bitmap.Assign(chrtsensitivity.TeeCreateBitmap
          (clWebPURPLE, rect(0, 0, imgsensgraph.Width, imgsensgraph.Height)));
      imagemax.Picture.Bitmap.Assign
        (chrtsensitivity.TeeCreateBitmap(clWebPURPLE, rect
            (0, 0, imagemax.Width, imagemax.Height)));



  end;
end;

procedure TfrmResultados.IWTimer1Timer(Sender: TObject);
var
  i, k: integer;
begin

  if TIWUserSession(WebApplication.Data).Start = True then
  sensitivityA;


  if TIWUserSession(WebApplication.Data).Finish = True then
  begin

      // para  escolha
      grdcrivarMConseq.Cell[0, 0].Text := 'Criterion';
      grdcrivarMConseq.ColumnCount := 3;
      grdcrivarMConseq.Cell[0, 2].Text := 'Upper Bound';
      grdcrivarMConseq.Cell[0, 1].Text := 'Lower Bound';

      grdcrivarMConseq.Cell[0,0].Font.Style := [fsbold];
      grdcrivarMConseq.Cell[0,1].Font.Style := [fsbold];
      grdcrivarMConseq.Cell[0,2].Font.Style := [fsbold];

      grdcrivarWeight.Cell[0, 0].Text := 'Criterion';
      grdcrivarWeight.ColumnCount := 3;
      grdcrivarWeight.Cell[0, 2].Text := 'Upper Bound';
      grdcrivarWeight.Cell[0, 1].Text := 'Lower Bound';

      grdcrivarWeight.Cell[0,0].Font.Style := [fsbold];
      grdcrivarWeight.Cell[0,1].Font.Style := [fsbold];
      grdcrivarWeight.Cell[0,2].Font.Style := [fsbold];

    k := 1;
    with TIWUserSession(WebApplication.Data) do
    begin
      IWTimer1.Enabled := false;
      //btnasok.Enabled := True;
      btnstartAS.Enabled := True;
      //btnBackSAV.Enabled := True;
      { for i := 0 to Rlevels - 1 do
        cmbxgroupindex.Items.Add('Group ' + IntToStr(i + 1));
        // cmbxgroupindex.ItemIndex := 0; }

      for i := 0 to NCrit - 1 do
        if CriSelMConseq[i] then
        begin
          INC(k);
          grdcrivarMConseq.RowCount := k;
          grdcrivarMConseq.Cell[k - 1, 0].Text := ObjSelecionados[i];
          grdcrivarMConseq.Cell[k - 1, 2].Text := '+' + FloatToStr(pmaxMConseq[i]) + '%';
          grdcrivarMConseq.Cell[k - 1, 1].Text := '-' + FloatToStr(pminMConseq[i]) + '%';
        end;

      k := 1;
      for i := 0 to NCrit - 1 do
        if CriSelWeight[i] then
        begin
          INC(k);
          grdcrivarWeight.RowCount := k;
          grdcrivarWeight.Cell[k - 1, 0].Text := ObjSelecionados[i];
          grdcrivarWeight.Cell[k - 1, 2].Text := '+' + FloatToStr(pmaxWeight[i]) + '%';
          grdcrivarWeight.Cell[k - 1, 1].Text := '-' + FloatToStr(pminWeight[i]) + '%';
        end;

      //Preencher_MatDadosAS;
    end; // with



    rgnsav.Visible := false;
    IWRgnAS.Visible := false;
    IWRgnanalisedesensibilidade.Visible := True;

    //cmbbxAS.Items.Clear;

  end;
end;

procedure TfrmResultados.lnkBackASClick(Sender: TObject);
begin
DestruirElementosAS;
rgnsav.Visible := False;
rgnresultados.Visible := True;
end;

procedure TfrmResultados.lstbxcriterioROCClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to TIWUserSession(WebApplication.Data).Ncrit - 1 do
  begin
    if TIWUserSession(WebApplication.Data).vetorGrafico[i] <> 2 then
    begin
      if i = lstbxcriterioROC.ItemIndex then
        TIWUserSession(WebApplication.Data).vetorGrafico[i] := 1
      else
        TIWUserSession(WebApplication.Data).vetorGrafico[i] := 0;
    end;
  end;
  gerar_grafico_overall;
end;

procedure TfrmResultados.cmbbxPlanosChange(Sender: TObject);
var
i,j : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
for i := 0 to Length(ObjSelecionados) - 1 do
begin
  grddesempenho.cell[1, i].text := FormatFloat('0.0000',PerformancePopulacao[cmbbxPlanos.ItemIndex, i]);
  grddesempenho.Cell[1, i].Font.FontName := 'verdana';
end;
for i := 0 to Length(populacao[0]) - 1 do
begin
grdplanos.Cell[i+1,0].Text := inttostr(i+1);
for j := 0 to Length(populacao[0,0]) - 1 do
grdplanos.Cell[i+1,j+1].Text := floattostr(populacao[cmbbxPlanos.ItemIndex,i,j]);
end;

end;

end;

procedure TfrmResultados.ConsequenceMatrixClick(Sender: TObject);
var
  ExcelApplication: variant;
  Sheet1, Sheet2, Sheet3: variant;
  i, j, z, k, t, counter1, counter2: integer;
  Matriz: Array of Array of Real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin

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



  end;
end;


procedure TfrmResultados.DetailedPlansClick(Sender: TObject);
var
i, counter2, k, z : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
    grdplans.columns.Clear;
    counter2 := -1;
    grdplans.TotalRows := 3;
    for I := 0 to 8 do
    begin
    grdplans.columns.Add;
    end;
    for I := 0 to Length(Populacao) - 1 do
    begin
    Inc(Counter2);
    grdplans.TotalRows := 2+ Counter2;
    grdplans.Cells[0,Counter2+1] := 'Plan ' + inttostr(i+1) + ':';
    //sheet2.Cells[Counter2+2, 1] := 'Plan ' + inttostr(i+1) + ':';
    for k := 0 to Length(Populacao[i]) - 1 do
    begin
    Inc(Counter2);
    grdplans.TotalRows := 2+ Counter2;
    //grdplans.Cells[0,Counter2+1] := inttostr(z+1);
      for z := 0 to Length(Populacao[i, k]) - 1 do
        begin
       grdplans.Cells[Z + 1,Counter2+1] := floatToStr(Populacao[i, k, z]);
       //sheet2.Cells[Counter2+2,z+2] := floatToStr(Populacao[i, k, z]);
        end;
    end;
    end;
   grdplans.Cells[0,0] := 'Plan';
   grdplans.Cells[1,0] := 'Regular Production';
   grdplans.Cells[2,0] := 'Overtime Production';
   grdplans.Cells[3,0] := 'Subcontrating Production';
   grdplans.Cells[4,0] := 'Stock';
   grdplans.Cells[5,0] := 'Hiring';
   grdplans.Cells[6,0] := 'Resignation';
   grdplans.Cells[7,0] := 'Total Employees';
TIWAdvWebGridExcelIO2.AdvWebGrid := grdplans;
  TIWAdvWebGridExcelIO2.XLSExport('Plans_' + problema + '.xls', WebApplication);
  //Inherited;
end;
end;

procedure TfrmResultados.rgnimagemaxCreate(Sender: TObject);
begin
rgnimagemax.Top := 169;
  rgnimagemax.left := 48;
end;

procedure TfrmResultados.rgnplanosCreate(Sender: TObject);
var
i : integer;
begin
rgnplanos.Align := AlClient;
With TIWUserSession(WebApplication.Data) do
begin
for i := 0 to Length(populacao) - 1 do
cmbbxPlanos.Items.Add('Plan ' + inttostr(i+1));
grddesempenho.clear;
grddesempenho.rowcount := 2;
grddesempenho.columncount := length(ObjSelecionados);

grdplanos.clear;
grdplanos.rowcount := Length(populacao[0])+1;
grdplanos.columncount := 8;
grdplanos.Cell[0,0].Text := 'Period';
grdplanos.Cell[0,1].Text := 'Regular Production';
grdplanos.Cell[0,2].Text := 'Overtime';
grdplanos.Cell[0,3].Text := 'Subcontrating';
grdplanos.Cell[0,4].Text := 'Stock';
grdplanos.Cell[0,5].Text := 'Hiring';
grdplanos.Cell[0,6].Text := 'Resignation';
grdplanos.Cell[0,7].Text := 'Total Employees';
for I := 0 to 7 do
begin
  grdplanos.cell[0, i].BGColor := clWebWHITESMOKE;
  grdplanos.cell[0, i].wrap := True;
  grdplanos.cell[0, i].Alignment := taCenter;
  grdplanos.Cell[0, i].Font.FontName := 'verdana';
  grdplanos.Cell[0, i].Font.Style := [fsbold];

end;
for i := 0 to Length(ObjSelecionados) - 1 do
begin
  grddesempenho.cell[0, i].text := ObjSelecionados[i];
  grddesempenho.cell[0, i].wrap := True;
  grddesempenho.cell[0, i].Alignment := taCenter;
  grddesempenho.Cell[0, i].Font.FontName := 'verdana';
  grddesempenho.Cell[0, i].Font.Style := [fsbold];
end;
end;
cmbbxplanos.ItemIndex := 0;
cmbbxPlanosChange(Self);
end;

procedure TfrmResultados.rgnresultadosCreate(Sender: TObject);
begin
rgnresultados.Align := AlClient;
end;

procedure TfrmResultados.rgnsavCreate(Sender: TObject);
begin
rgnsav.align := AlClient;
end;

procedure TfrmResultados.rgnsmarterCreate(Sender: TObject);
begin
rgnsmarter.Align := AlClient;
end;

initialization
  TfrmResultados.SetAsMainForm;

end.
