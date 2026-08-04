// vybe-test: pascal/pascal_class_properties_accessors/test_property_index_specifier
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TDataHolder = class
  private FVal1, FVal2: Integer;
  private function GetVal(index: Integer): Integer;
  private procedure SetVal(index: Integer; value: Integer);
  public property Val1: Integer index 1 read GetVal write SetVal;
  public property Val2: Integer index 2 read GetVal write SetVal;
end;
function TDataHolder.GetVal(index: Integer): Integer;
begin
  if index = 1 then Result := FVal1 else Result := FVal2;
end;
procedure TDataHolder.SetVal(index: Integer; value: Integer);
begin
  if index = 1 then FVal1 := value else FVal2 := value;
end;
var h: TDataHolder;
begin
  h := TDataHolder.Create;
  h.Val1 := 10;
  h.Val2 := 20;
  WriteLn(h.Val1);
  WriteLn(h.Val2);
  h.Free;
end.
