// vybe-test: pascal/pascal_class_properties_accessors/test_property_setter_validation_side_effect
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TTemperature = class
  private FCelsius: Integer;
  private procedure SetCelsius(v: Integer);
  public property Celsius: Integer read FCelsius write SetCelsius;
end;
procedure TTemperature.SetCelsius(v: Integer);
begin
  if v < -273 then FCelsius := -273
  else FCelsius := v;
end;
var t: TTemperature;
begin
  t := TTemperature.Create;
  t.Celsius := -300;
  WriteLn(t.Celsius);
  t.Celsius := 25;
  WriteLn(t.Celsius);
  t.Free;
end.
