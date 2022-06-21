unit uPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, IniFiles,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.DB2,
  FireDAC.Phys.DB2Def, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.VCLUI.Wait, FireDAC.VCLUI.Error, Vcl.StdCtrls,
  FireDAC.Comp.UI, FireDAC.Phys.ODBCBase, FireDAC.Comp.Client, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Phys.ASA, FireDAC.Phys.ASADef, Vcl.Grids,
  Vcl.DBGrids, uConexao, Vcl.ExtCtrls, Vcl.Menus, Vcl.AppEvnts;

type
  TForm1 = class(TForm)
    FDQuery1: TFDQuery;
    Transaction: TFDTransaction;
    FDPhysDB2DriverLink1: TFDPhysDB2DriverLink;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDGUIxErrorDialog1: TFDGUIxErrorDialog;
    Button1: TButton;
    Button2: TButton;
    DataSource1: TDataSource;
    Button3: TButton;
    Timer1: TTimer;
    lblMSg: TLabel;
    Label1: TLabel;
    Conexao: TFDConnection;
    TrayIcon1: TTrayIcon;
    PopupMenu1: TPopupMenu;
    Abrir1: TMenuItem;
    Abrir2: TMenuItem;
    Sair1: TMenuItem;
    ApplicationEvents1: TApplicationEvents;
    lblStatus: TLabel;
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure Abrir1Click(Sender: TObject);
    procedure ApplicationEvents1Minimize(Sender: TObject);
  private
    { Private declarations }
    procedure ExportSWEDA;
    procedure conect;

  public
    { Public declarations }
    arqINI: TIniFile;

  end;

var
  Form1: TForm1;
  TempoSeg, Servidor, Banco, Porta, Usuario, Senha, DriveID, PATH: String;
  cont, seg: Integer;
  msg, qtdSeg: String;

implementation

{$R *.dfm}

procedure TForm1.conect;
begin
  arqINI := TIniFile.Create(ExtractFilePath(Application.ExeName) +
    'Config.ini');
  try
    with Conexao do
    begin
      Params.Clear;
      Connected := false;
      LoginPrompt := false;
      Params.Values['DriverID'] := DriveID;
      Params.Values['Server'] := Servidor;
      Params.Values['Port'] := Porta;
      Params.Values['Database'] := Banco;
      Params.Values['User_name'] := Usuario;
      Params.Values['Password'] := Senha;
      Connected := true;
      if Connected = true then
      begin
        // ShowMessage('Conectado com SUCESSO!');
        // sleep(1000);
        // FDQuery1.Open();
        // Label1.Caption := 'Conectado com SUCESSO!';
      end;
    end;
  except
    on E: Exception do
      ShowMessage('Erro ao carregar parâmetros de conexão!'#13#10 + E.Message);
  end;
end;

procedure TForm1.ExportSWEDA;
var
  Stream: TFileStream;
  i: Integer;
  OutLine, f: string;
  sTemp, s: string;
  { Código da net }
begin
  lblStatus.Caption := 'Aguarde! Realizando a exportação de preço.';
  FDQuery1.Open();
  Stream := TFileStream.Create(PATH, fmCreate);
  try
    Application.ProcessMessages;
    FDQuery1.First;
    while not FDQuery1.Eof do
    begin

      s := '';
      OutLine := '';
      for i := 0 to FDQuery1.FieldCount - 1 do
      begin
        sTemp := FDQuery1.Fields[i].AsString;
        // Special handling to sTemp here
        OutLine := OutLine + sTemp + ',';
      end;
      // Remove final unnecessary ','
      SetLength(OutLine, Length(OutLine) - 1);
      // Write line to file
      Stream.Write(OutLine[1], Length(OutLine) * SizeOf(Char));
      // Write line ending
      Stream.Write(sLineBreak, Length(sLineBreak));
      FDQuery1.Next;
      Application.ProcessMessages;
    end;
  finally
    Stream.Free; // Saves the file
  end;
  // ShowMessage('Exportado com sucesso!');
  TrayIcon1.ShowBalloonHint();
  TrayIcon1.Visible := true;
  TrayIcon1.BalloonHint := 'Exportado SWEED.csv com sucesso!';
  TrayIcon1.ShowBalloonHint();
  lblStatus.Caption := '';
  // FDQuery1.Close;
  // FDQuery1.Cancel;
end;

procedure TForm1.Abrir1Click(Sender: TObject);
begin
  TrayIcon1.Visible := false;
  Show();
  WindowState := wsNormal;
  Application.BringToFront();
end;

procedure TForm1.ApplicationEvents1Minimize(Sender: TObject);
begin
  Self.Hide();
  TrayIcon1.ShowBalloonHint();
  Self.WindowState := wsMinimized;
  TrayIcon1.BalloonHint := 'Aplicação executando em segundo plano.';
  TrayIcon1.Visible := true;
  TrayIcon1.Animate := true;
  TrayIcon1.ShowBalloonHint;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  FDQuery1.Open();
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  ExportSWEDA;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // Carrega as informações do arquivo INI nos campos
  arqINI := TIniFile.Create(ExtractFilePath(Application.ExeName) +
    'Config.ini');
  if not(FileExists(ExtractFilePath(Application.ExeName) + 'Config.ini')) then
  begin
    TempoSeg := '60'; // 60 = 1minuto
    arqINI.WriteString('CONFIG', 'Tempo', TempoSeg);
    arqINI.WriteString('ECD', 'Local', 'C:\ECD1200\SWEDA.csv');
    arqINI.WriteString('Configuração', '==============================', '');

    arqINI.WriteString('Conexao', 'DriverID', DriveID);
    arqINI.WriteString('Conexao', 'Servidor', Servidor);
    arqINI.WriteString('Conexao', 'Porta', Porta);
    arqINI.WriteString('Conexao', 'Database', Banco);
    arqINI.WriteString('Conexao', 'Usuario', Usuario);
    arqINI.WriteString('Conexao', 'Senha', Senha);
    arqINI.Free;

  end;

end;

procedure TForm1.FormShow(Sender: TObject);
begin
  try
    arqINI := TIniFile.Create(ExtractFilePath(Application.ExeName) +
      'Config.ini');

    TempoSeg := arqINI.ReadString('CONFIG', 'Tempo', '');
    lblMSg.Caption := 'PRÓX. EXPORTAÇÃO EM: ' + (TempoSeg) + ' seg.';
    PATH := arqINI.ReadString('ECD', 'Local', '');

    DriveID := arqINI.ReadString('Conexao', 'DriverID', '');
    Servidor := arqINI.ReadString('Conexao', 'Servidor', '');
    Porta := arqINI.ReadString('Conexao', 'Porta', '');
    Banco := arqINI.ReadString('Conexao', 'Database', '');
    Usuario := arqINI.ReadString('Conexao', 'Usuario', '');
    Senha := arqINI.ReadString('Conexao', 'Senha', '');
    arqINI.Free;
    conect;
  except
    on E: Exception do
      ShowMessage('Erro ao carregar parâmetros de conexão!'#13#10#13#10 +
        E.Message);
  end;
end;

procedure TForm1.Sair1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin

  arqINI := TIniFile.Create(ExtractFilePath(Application.ExeName) +
    'Config.ini');

  seg := StrToInt(TempoSeg);
  cont := seg - 1;
  TempoSeg := IntToStr(cont);

  lblMSg.Caption := 'PRÓX. EXPORTAÇÃO EM: ' + IntToStr(cont) + ' seg.';

  if cont < 1 then
  begin
    // FDQuery1.Open();
    lblStatus.Caption := 'Aguarde! Realizando a exportação de preço.';
    TempoSeg := arqINI.ReadString('CONFIG', 'Tempo', '');
    try
      ExportSWEDA; // Função de exporta para Excel;
    except
      // ShowMessage('Falha ao excluir os arquivos!');
      Exit;
      lblStatus.Visible := false;
      Timer1.Enabled := false;
      arqINI.Free;
    end;
  end;
end;

end.
