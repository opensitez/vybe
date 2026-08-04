// vybe-test: pascal/properties/test_property_class_level
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TApp = class
  private
    class var FVersion: string;
    class function GetVersion: string; static;
    class procedure SetVersion(v: string); static;
  public
    class property Version: string read GetVersion write SetVersion;
  end;

class function TApp.GetVersion: string;
begin
  Result := FVersion;
end;

class procedure TApp.SetVersion(v: string);
begin
  FVersion := v;
end;

begin
  TApp.Version := '1.0.0';
  WriteLn(TApp.Version);
  TApp.Version := '2.0.0';
  WriteLn(TApp.Version);
end.
