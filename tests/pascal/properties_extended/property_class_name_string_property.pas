// vybe-test: pascal/properties_extended/property_class_name_string_property
// origin: languages/pascal/tests/pascal/test_properties_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TMeta=class strict private class var FTag:string; class function GetTag:string; class procedure SetTag(const s:string); public class property Tag:string read GetTag write SetTag; end; class function TMeta.GetTag:string; begin Result:=FTag; end; class procedure TMeta.SetTag(const s:string); begin FTag:=s; end; begin TMeta.Tag:='pascal'; WriteLn(TMeta.Tag); end.
