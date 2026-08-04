// vybe-test: pascal/procedural_types/procedural_type_filter_with_callback
// origin: languages/pascal/tests/pascal/test_procedural_types.rs
program T;
{$mode delphi}
uses SysUtils; function CountIf(a: array of Integer; pred: function(x: Integer): Boolean): Integer; var i: Integer; begin Result:=0; for i:=0 to High(a) do if pred(a[i]) then Result:=Result+1; end; begin WriteLn(CountIf([1,2,3,4], function(x: Integer): Boolean begin Result:=x>2; end)); end.
