// vybe-test: pascal/properties/test_property_in_loop
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TCounter = class
  private
    FValue: Integer;
    procedure SetValue(v: Integer);
    function GetValue: Integer;
  public
    property Value: Integer read GetValue write SetValue;
    procedure Tick;
  end;

procedure TCounter.SetValue(v: Integer);
begin
  if v >= 0 then FValue := v;
end;

function TCounter.GetValue: Integer;
begin
  Result := FValue;
end;

procedure TCounter.Tick;
begin
  FValue := FValue + 1;
end;

var
  c: TCounter;
  i: Integer;
begin
  c := TCounter.Create;
  c.Value := 0;
  for i := 1 to 5 do
    c.Tick;
  WriteLn(c.Value);
end.
