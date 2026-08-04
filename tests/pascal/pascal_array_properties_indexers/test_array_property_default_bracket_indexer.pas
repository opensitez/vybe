// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_default_bracket_indexer
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TIntList = class
  private FData: array[0..4] of Integer;
  private function GetItem(index: Integer): Integer;
  private procedure SetItem(index, val: Integer);
  public property Items[index: Integer]: Integer read GetItem write SetItem; default;
end;
function TIntList.GetItem(index: Integer): Integer; begin Result := FData[index]; end;
procedure TIntList.SetItem(index, val: Integer); begin FData[index] := val; end;
var list: TIntList;
begin
  list := TIntList.Create;
  list[0] := 10;
  list[1] := 20;
  WriteLn(list[0]);
  WriteLn(list[1]);
  list.Free;
end.
