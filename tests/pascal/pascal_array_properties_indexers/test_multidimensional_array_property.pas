// vybe-test: pascal/pascal_array_properties_indexers/test_multidimensional_array_property
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TGrid = class
  private FCells: array[0..2, 0..2] of Integer;
  private function GetCell(row, col: Integer): Integer;
  private procedure SetCell(row, col, val: Integer);
  public property Cells[row, col: Integer]: Integer read GetCell write SetCell; default;
end;
function TGrid.GetCell(row, col: Integer): Integer; begin Result := FCells[row, col]; end;
procedure TGrid.SetCell(row, col, val: Integer); begin FCells[row, col] := val; end;
var g: TGrid;
begin
  g := TGrid.Create;
  g[1, 2] := 99;
  WriteLn(g[1, 2]);
  g.Free;
end.
