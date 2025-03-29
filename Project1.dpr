program Project1;

uses
  Forms,
  JvGnugettext,
  Unit1 in 'Unit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Importador Lançamentos Contábil';
  UseLanguage('pt_BR');
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
