// vybe-test: pascal/procedural_types/procedural_type_map_with_callback
// origin: languages/pascal/tests/pascal/test_procedural_types.rs
program T;
{$mode delphi}
uses SysUtils; function Apply(a: array of Integer; fn: function(x: Integer): Integer): Integer; begin Result:=fn(a[0])+fn(a[1]); end; begin WriteLn(Apply([3,4], function(x: Integer): Integer begin Result:=x*x; end)); end.
