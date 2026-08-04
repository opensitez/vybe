// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_bounds_checking_setter
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBoundedArray = class
  private FItems: array[0..2] of Integer;
  private procedure SetItem(i: Integer; val: Integer);
  private function GetItem(i: Integer): Integer;
  public property Items[i: Integer]: Integer read GetItem write SetItem; default;
end;
procedure TBoundedArray.SetItem(i: Integer; val: Integer);
begin
  if (i >= 0) and (i <= 2) then FItems[i] := val;
end;
function TBoundedArray.GetItem(i: Integer): Integer;
begin
  if (i >= 0) and (i <= 2) then Result := FItems[i] else Result := -1;
end;
var ba: TBoundedArray;
begin
  ba := TBoundedArray.Create;
  ba[0] := 42;
  ba[5] := 99;
  WriteLn(ba[0]);
  WriteLn(ba[5]);
  ba.Free;
end.
