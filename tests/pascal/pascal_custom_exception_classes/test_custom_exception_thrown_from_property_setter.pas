// vybe-test: pascal/pascal_custom_exception_classes/test_custom_exception_thrown_from_property_setter
// origin: languages/pascal/tests/pascal/test_pascal_custom_exception_classes.rs
program Test;
{$mode delphi}
uses SysUtils;
type EInvalidAge = class(Exception);
type TPerson = class
  private FAge: Integer;
  private procedure SetAge(v: Integer);
  public property Age: Integer read FAge write SetAge;
end;
procedure TPerson.SetAge(v: Integer);
begin
  if v < 0 then raise EInvalidAge.Create('AgeCannotBeNegative');
  FAge := v;
end;
var p: TPerson;
begin
  p := TPerson.Create;
  try
    p.Age := -10;
  except
    on E: EInvalidAge do WriteLn(E.Message);
  end;
  p.Free;
end.
