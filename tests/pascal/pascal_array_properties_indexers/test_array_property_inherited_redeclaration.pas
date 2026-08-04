// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_inherited_redeclaration
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBase = class
  protected FValues: array[0..2] of Integer;
  protected function GetVal(i: Integer): Integer;
  public property Values[i: Integer]: Integer read GetVal; default;
end;
type TDerived = class(TBase)
  protected procedure SetVal(i, v: Integer);
  public property Values[i: Integer]: Integer read GetVal write SetVal; default;
end;
function TBase.GetVal(i: Integer): Integer; begin Result := FValues[i]; end;
procedure TDerived.SetVal(i, v: Integer); begin FValues[i] := v; end;
var d: TDerived;
begin
  d := TDerived.Create;
  d[1] := 77;
  WriteLn(d[1]);
  d.Free;
end.
