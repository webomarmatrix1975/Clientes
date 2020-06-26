program Clientes;

uses
  Vcl.Forms,
  UClientes in 'UClientes.pas' {FrmClientes};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmClientes, FrmClientes);
  Application.Run;
end.
