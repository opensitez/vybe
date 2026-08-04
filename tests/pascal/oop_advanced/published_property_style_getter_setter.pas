// vybe-test: pascal/oop_advanced/published_property_style_getter_setter
// origin: languages/pascal/tests/pascal/test_oop_advanced.rs
program T;
{$mode delphi}
uses SysUtils;
type TProp=class
  private FV:Integer;
  public function GetV:Integer; procedure SetV(v:Integer); property Value:Integer read GetV write SetV;
end;
function TProp.GetV:Integer; begin Result:=FV; end;
procedure TProp.SetV(v:Integer); begin FV:=v; end;
var p:TProp; begin p:=TProp.Create; p.Value:=9; WriteLn(p.Value); p.Free; end.
