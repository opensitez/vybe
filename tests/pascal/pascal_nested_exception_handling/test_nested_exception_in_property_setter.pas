// vybe-test: pascal/pascal_nested_exception_handling/test_nested_exception_in_property_setter
// origin: languages/pascal/tests/pascal/test_pascal_nested_exception_handling.rs
program Test;
{$mode delphi}
uses SysUtils;
type TTestProp = class
  private procedure SetVal(v: Integer);
  public property Val: Integer write SetVal;
end;
procedure TTestProp.SetVal(v: Integer);
begin
  if v < 0 then raise Exception.Create('NegativeValueNotAllowed');
end;
var t: TTestProp;
begin
  t := TTestProp.Create;
  try
    t.Val := -5;
  except
    on E: Exception do WriteLn('CaughtSetter:' + E.Message);
  end;
  t.Free;
end.
