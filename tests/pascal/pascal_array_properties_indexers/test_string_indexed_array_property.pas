// vybe-test: pascal/pascal_array_properties_indexers/test_string_indexed_array_property
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TMapHolder = class
  private FKeys: array[0..2] of String;
  private FValues: array[0..2] of String;
  private function GetVal(key: String): String;
  private procedure SetVal(key, val: String);
  public constructor Create;
  public property Values[key: String]: String read GetVal write SetVal; default;
end;
constructor TMapHolder.Create;
begin
  FKeys[0] := 'host'; FKeys[1] := 'port';
end;
function TMapHolder.GetVal(key: String): String;
begin
  if key = 'host' then Result := FValues[0]
  else if key = 'port' then Result := FValues[1]
  else Result := '';
end;
procedure TMapHolder.SetVal(key, val: String);
begin
  if key = 'host' then FValues[0] := val
  else if key = 'port' then FValues[1] := val;
end;
var map: TMapHolder;
begin
  map := TMapHolder.Create;
  map['host'] := 'localhost';
  map['port'] := '8080';
  WriteLn(map['host'] + ':' + map['port']);
  map.Free;
end.
