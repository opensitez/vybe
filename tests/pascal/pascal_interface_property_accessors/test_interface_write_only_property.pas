// vybe-test: pascal/pascal_interface_property_accessors/test_interface_write_only_property
// origin: languages/pascal/tests/pascal/test_pascal_interface_property_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type IWriteOnlyLogger = interface
  ['{55555555-6666-7777-8888-999999999999}']
  procedure SetLogMsg(const msg: String);
  property LogMsg: String write SetLogMsg;
end;

type TLoggerImpl = class(TInterfacedObject, IWriteOnlyLogger)
  public procedure SetLogMsg(const msg: String);
end;
procedure TLoggerImpl.SetLogMsg(const msg: String);
begin
  WriteLn('Logged:' + msg);
end;

var l: IWriteOnlyLogger;
begin
  l := TLoggerImpl.Create;
  l.LogMsg := 'InterfaceWriteOnly';
end.
