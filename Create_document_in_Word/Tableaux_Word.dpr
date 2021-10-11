program Tableaux_Word;

uses
  Forms,
  Tableaux_Word_u in 'Tableaux_Word_u.pas' {Form1},
  DriveOleWord in 'DriveOleWord.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
