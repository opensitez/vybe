// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_in_nested_loops
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TMatrix2x2 = class
  private FData: array[0..1, 0..1] of Integer;
  private function GetV(r, c: Integer): Integer;
  private procedure SetV(r, c, v: Integer);
  public property Cells[r, c: Integer]: Integer read GetV write SetV; default;
end;
function TMatrix2x2.GetV(r, c: Integer): Integer; begin Result := FData[r, c]; end;
procedure TMatrix2x2.SetV(r, c, v: Integer); begin FData[r, c] := v; end;
var m: TMatrix2x2; r, c, sum: Integer;
begin
  m := TMatrix2x2.Create;
  m[0, 0] := 1; m[0, 1] := 2;
  m[1, 0] := 3; m[1, 1] := 4;
  sum := 0;
  for r := 0 to 1 do
    for c := 0 to 1 do
      sum := sum + m[r, c];
  WriteLn(sum);
  m.Free;
end.
