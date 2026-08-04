// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_float_values
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TFloatList = class
  private FFloats: array[0..2] of Real;
  private function GetF(i: Integer): Real;
  private procedure SetF(i: Integer; v: Real);
  public property Floats[i: Integer]: Real read GetF write SetF; default;
end;
function TFloatList.GetF(i: Integer): Real; begin Result := FFloats[i]; end;
procedure TFloatList.SetF(i: Integer; v: Real); begin FFloats[i] := v; end;
var fl: TFloatList;
begin
  fl := TFloatList.Create;
  fl[0] := 1.1; fl[1] := 2.2;
  WriteLn(fl[0] + fl[1]);
  fl.Free;
end.
