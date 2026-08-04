// vybe-test: pascal/properties_extended/property_class_static_read_write
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TApp=class strict private class var FName:string; class function GetName:string; class procedure SetName(const v:string); public class property AppName:string read GetName write SetName; end; class function TApp.GetName:string; begin Result:=FName; end; class procedure TApp.SetName(const v:string); begin FName:=v; end; begin TApp.AppName:='vybe'; WriteLn(TApp.AppName); end.
