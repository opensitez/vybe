// vybe-test: pascal/pascal_array_properties_indexers/test_enum_indexed_array_property
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TColor = (cRed, cGreen, cBlue);
type TPalette = class
  private FColors: array[TColor] of String;
  private function GetHex(c: TColor): String;
  private procedure SetHex(c: TColor; hex: String);
  public property ColorHex[c: TColor]: String read GetHex write SetHex; default;
end;
function TPalette.GetHex(c: TColor): String; begin Result := FColors[c]; end;
procedure TPalette.SetHex(c: TColor; hex: String); begin FColors[c] := hex; end;
var p: TPalette;
begin
  p := TPalette.Create;
  p[cRed] := '#FF0000';
  p[cGreen] := '#00FF00';
  WriteLn(p[cRed]);
  WriteLn(p[cGreen]);
  p.Free;
end.
