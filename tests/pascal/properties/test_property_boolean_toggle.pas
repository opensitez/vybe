// vybe-test: pascal/properties/test_property_boolean_toggle
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TSwitch = class
  private
    FOn: Boolean;
    procedure SetOn(v: Boolean);
  public
    property IsOn: Boolean read FOn write SetOn;
    procedure Toggle;
  end;

procedure TSwitch.SetOn(v: Boolean);
begin
  FOn := v;
end;

procedure TSwitch.Toggle;
begin
  FOn := not FOn;
end;

var
  sw: TSwitch;
begin
  sw := TSwitch.Create;
  sw.IsOn := false;
  WriteLn(sw.IsOn);
  sw.Toggle;
  WriteLn(sw.IsOn);
  sw.Toggle;
  WriteLn(sw.IsOn);
end.
