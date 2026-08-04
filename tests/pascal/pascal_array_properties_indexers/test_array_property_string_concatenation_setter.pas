// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_string_concatenation_setter
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TLogBuffer = class
  private FLogs: array[0..2] of String;
  private procedure AppendLog(i: Integer; msg: String);
  private function GetLog(i: Integer): String;
  public property Logs[i: Integer]: String read GetLog write AppendLog; default;
end;
procedure TLogBuffer.AppendLog(i: Integer; msg: String);
begin
  FLogs[i] := FLogs[i] + msg;
end;
function TLogBuffer.GetLog(i: Integer): String; begin Result := FLogs[i]; end;
var lb: TLogBuffer;
begin
  lb := TLogBuffer.Create;
  lb[0] := 'Line1:';
  lb[0] := ' OK';
  WriteLn(lb[0]);
  lb.Free;
end.
