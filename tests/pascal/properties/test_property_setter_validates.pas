// vybe-test: pascal/properties/test_property_setter_validates
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TRange = class
  private
    FValue: Integer;
    FMin, FMax: Integer;
    procedure SetValue(v: Integer);
  public
    constructor Create(mn, mx: Integer);
    property Value: Integer read FValue write SetValue;
  end;

constructor TRange.Create(mn, mx: Integer);
begin
  inherited Create;
  FMin := mn;
  FMax := mx;
  FValue := mn;
end;

procedure TRange.SetValue(v: Integer);
begin
  if v < FMin then FValue := FMin
  else if v > FMax then FValue := FMax
  else FValue := v;
end;

var
  r: TRange;
begin
  r := TRange.Create(0, 100);
  r.Value := 50;
  WriteLn(r.Value);
  r.Value := -10;
  WriteLn(r.Value);
  r.Value := 200;
  WriteLn(r.Value);
end.
