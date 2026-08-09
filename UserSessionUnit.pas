unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to 
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, SysUtils, Classes, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection, IWCompEdit,
  IWCompLabel, IWGrids, series, IWRegion, IWCompTabControl, IWCompListbox;

type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery1: TZQuery;
    ZConnection1: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
    MatrizFinal: Array of Array of String;
  Demanda: Array of Real;
  Limites: Array of Real;
  Custos: Array of Real;
  Periodo : integer;
  ObjSelecionados : array of string;
  ParametrosFuncaoObjetivo: Array of Array of Real;
  Populacao, Population, NovaGeracao: Array of Array of Array of Real;
  // MatrizAvaliacao
  PopulationPerformance, PerformancePopulacao, PerformanceNormalizada: Array of Array of Real;
  PopulationRanking, OrdemPopulacao: array of array of integer;
  PopulationDistance, DistanciaPopulacional: Array of Real;
  NinicialFunc, Nger, G1, G2: integer;
  DemandaMedia: Real;
  N1, analyst : integer;
  Problema, Id_User : string;
  primeiroacesso : boolean;
  email, Id_Email : string;
  senha, validation: string;
  nameproblem, idproblemasuser, datainicioproblema, datafinalproblema : array of string;
  ordemsel : array of integer;
  auxiliar : integer;
  datainicio : string;
  retomada : boolean;
  VetorGrafico, indexROC : array of integer;
  maior, menor, valor_razao, BConseqOrd, WConseqOrd, peso : array of real;
  Nalt, Ncrit, melhorplano, melhorplanoAS : integer;
  Reinicia : integer;
  edtmaxMConseq, edtminMConseq : array of tiwEdit;
  cmbbxmaxMConseq, cmbbxminMConseq : array of Tiwcombobox;
  edtmaxWeight, edtminWeight : array of tiwEdit;
  lblmaxMconseq, lblminMconseq : array of TiwLabel;
  lblmaxWeight, lblminWeight : array of TiwLabel;
  lblvariationmaxMConseq, lblvariationminMConseq : array of TiwLabel;
  lblVariationmaxWeight, lblVariationminWeight : array of TiwLabel;
  PmaxMConseq, PminMConseq : array of double;
  CriselMConseq : array of boolean;
  PmaxWeight, PminWeight : array of double;
  CriselWeight : array of boolean;
  Start, Finish : boolean;
  NumCases : integer;
  pesoAS : array of real;
  PerformancePopulacaoAS, PerformanceNormalizadaAS : array of array of real;
  RelatorioAS : array of integer;
  maiorAS, menorAS : array of real;
  idproblemapadrao : string;
  Id_Problem : string;
  problemapadrao : boolean;
  Actor, ActorDescription : array of string;
  ContextProblem : string;
  end;

implementation

uses Unit2, ServerController, untResultados;

{$R *.dfm}

end.