// vybe-test: pascal/pascal_class_properties_accessors/test_property_method_accessors
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TAccount = class
  private FBalance: Real;
  private function GetBalance: Real;
  private procedure SetBalance(v: Real);
  public property Balance: Real read GetBalance write SetBalance;
end;
function TAccount.GetBalance: Real; begin Result := FBalance; end;
procedure TAccount.SetBalance(v: Real); begin FBalance := v; end;
var acc: TAccount;
begin
  acc := TAccount.Create;
  acc.Balance := 250.50;
  WriteLn(acc.Balance);
  acc.Free;
end.
