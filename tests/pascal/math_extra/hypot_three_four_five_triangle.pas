// vybe-test: pascal/math_extra/hypot_three_four_five_triangle
// origin: languages/pascal/tests/pascal/test_math_extra.rs
program T;
{$mode delphi}
uses SysUtils;
function Hypot(a, b: Real): Real;
begin
  Result := Sqrt(a * a + b * b);
end;
begin
  WriteLn(Hypot(3.0, 4.0):0:0);
end.
