// vybe-test: pascal/procedural_types/procedural_type_compare_callback_sort_key
// origin: languages/pascal/tests/pascal/test_procedural_types.rs
program T;
{$mode delphi}
uses SysUtils; function PickMax(a,b: Integer; better: function(x,y: Integer): Boolean): Integer; begin if better(a,b) then Result:=a else Result:=b; end; begin WriteLn(PickMax(3,9, function(x,y: Integer): Boolean begin Result:=x>y; end)); end.
