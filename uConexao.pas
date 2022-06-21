unit uConexao;

interface

uses
  IniFiles, SysUtils, Forms, FireDAC.Comp.Client, Dialogs;

type
  TConexao = class
  private
    FPath: string;
    FServidor: string;
    FPorta: integer;
    FDatabase: string;
    FSenha: string;
    FUsuario: string;
    FDriver: string;
    FSecao: string;
  public
    property Path: string read FPath write FPath;
    property Servidor: string read FServidor write FServidor;
    property Porta: integer read FPorta write FPorta;
    property Database: string read FDatabase write FDatabase;
    property Senha: string read FSenha write FSenha;
    property Usuario: string read FUsuario write FUsuario;
    property Driver: string read FDriver write FDriver;
    property Secao: string read FSecao write FSecao;

    constructor Create(Path: string; Secao: string);

    procedure LeINI(); virtual;
    procedure GravaINI(Usuario, Senha, Servidor, Banco: string;
      Porta: integer); virtual;
    procedure Conectar(var Conexao: TFDConnection); virtual;
  end;

implementation

// Lê o arquivo INI e carrega os parâmetros no componente TFDConnection
procedure TConexao.Conectar(var Conexao: TFDConnection);
begin
  LeINI();

  try
    // Passa os parâmetros para o objeto Conexão
    Conexao.Connected := false;
    Conexao.LoginPrompt := false;
    Conexao.Params.Clear;
    Conexao.Params.Add('Server=' + Servidor);
    Conexao.Params.Add('User_name=' + Usuario);
    Conexao.Params.Add('Password=' + Senha);
    Conexao.Params.Add('Port=' + IntToStr(Porta));
    Conexao.Params.Add('Database=' + Database);
    Conexao.Params.Add('DriverID=' + Driver);
  Except
    on E: Exception do
      ShowMessage('Erro ao carregar parâmetros de conexão!'#13#10 + E.Message);
  end;
end;

// Método construtor recebe o caminho do INI e nome da Seção para Leitura
constructor TConexao.Create(Path: string; Secao: string);
var
  ArqIni: TIniFile;
begin
   ArqIni := TIniFile.Create(ExtractFilePath(Application.ExeName)+Path);
  if  not FileExists(ExtractFilePath(Application.ExeName)+Path) then
  begin
    Self.Path := Path;
    Self.Secao := Secao;

      GravaINI('teste', 'aa', 'teste', 'ate', StrToInt('1542'));
      //MessageDlg('Gravado com sucesso!', mtConfirmation, [mbOK], 0);
    end
    else
      raise Exception.Create
        ('Arquivo INI para configuração não encontrado.'#13#10'Aplicação será finalizada.');

end;

// Grava os parâmetros recebidos no arquivo INI
procedure TConexao.GravaINI(Usuario, Senha, Servidor, Banco: string;
  Porta: integer);
var
  ArqIni: TIniFile;
begin
  ArqIni := TIniFile.Create(ExtractFilePath(Application.ExeName)+Path);
  if  not FileExists(ExtractFilePath(Application.ExeName)+Path) then
  begin
  try
    ArqIni.WriteString(Secao, 'Usuario', Usuario);
    ArqIni.WriteString(Secao, 'Senha', Senha);
    ArqIni.WriteString(Secao, 'Database', Banco);
    ArqIni.WriteString(Secao, 'Servidor', Servidor);
    ArqIni.WriteInteger(Secao, 'Porta', Porta);
  finally
    ArqIni.Free;
  end;
end;
end;

// Lê os parâmetros do arquivo INI e atribui para os atributos
procedure TConexao.LeINI();
var
  ArqIni: TIniFile;
begin
  ArqIni := TIniFile.Create(Path);
  try
    Servidor := ArqIni.ReadString(Secao, 'Servidor', '');
    Porta := ArqIni.ReadInteger(Secao, 'Porta', 0);
    Database := ArqIni.ReadString(Secao, 'Database', '');
    Senha := ArqIni.ReadString(Secao, 'Senha', '');
    Usuario := ArqIni.ReadString(Secao, 'Usuario', '');
    Driver := ArqIni.ReadString(Secao, 'drivername', '');
  finally
    ArqIni.Free;
  end;
end;

end.
