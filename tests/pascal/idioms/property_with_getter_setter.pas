// vybe-test: pascal/idioms/property_with_getter_setter
// origin: languages/pascal/tests/pascal/test_idioms.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TTemperature = class
  private
    FCelsius: Real;
    function GetFahrenheit: Real;
    procedure SetFahrenheit(val: Real);
  public
    constructor Create;
    property Celsius: Real read FCelsius write FCelsius;
    property Fahrenheit: Real read GetFahrenheit write SetFahrenheit;
  end;

constructor TTemperature.Create; begin FCelsius := 0; end;
function TTemperature.GetFahrenheit: Real;
begin Result := FCelsius * 9 / 5 + 32; end;
procedure TTemperature.SetFahrenheit(val: Real);
begin FCelsius := (val - 32) * 5 / 9; end;

var t: TTemperature;
begin
  t := TTemperature.Create;
  t.Celsius := 100;
  WriteLn(t.Fahrenheit);
end.
