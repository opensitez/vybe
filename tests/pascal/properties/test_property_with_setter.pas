// vybe-test: pascal/properties/test_property_with_setter
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TPositive = class
  private
    FVal: Integer;
    procedure SetVal(v: Integer);
  public
    property Val: Integer read FVal write SetVal;
  end;

procedure TPositive.SetVal(v: Integer);
begin
  if v >= 0 then FVal := v
  else FVal := 0;
end;

var
  p: TPositive;
begin
  p := TPositive.Create;
  p.Val := 10;
  WriteLn(p.Val);
  p.Val := -5;
  WriteLn(p.Val);
end.
