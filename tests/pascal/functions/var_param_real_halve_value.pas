// vybe-test: pascal/functions/var_param_real_halve_value
// origin: languages/pascal/tests/pascal/test_functions.rs
program T;
{$mode delphi}
uses SysUtils;
procedure Halve(var x: Real);
begin
  x := x / 2.0;
end;
var r: Real;
begin
  r := 9.0;
  Halve(r);
  WriteLn(r:0:1);
end.
