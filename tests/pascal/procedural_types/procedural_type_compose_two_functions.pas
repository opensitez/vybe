// vybe-test: pascal/procedural_types/procedural_type_compose_two_functions
// origin: languages/pascal/tests/pascal/test_procedural_types.rs
program T;
{$mode delphi}
uses SysUtils; function Compose(f,g: function(x: Integer): Integer; v: Integer): Integer; begin Result:=f(g(v)); end; begin WriteLn(Compose(function(x: Integer): Integer begin Result:=x+1; end, function(x: Integer): Integer begin Result:=x*2; end, 3)); end.
