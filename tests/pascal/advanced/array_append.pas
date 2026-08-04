// vybe-test: pascal/advanced/array_append
// origin: languages/pascal/tests/pascal/test_advanced.rs
program T;
{$mode delphi}
uses SysUtils;
var a: array of Integer;
begin
  a := [1, 2];
  Append(a, 3);
  WriteLn(Length(a));
  WriteLn(a[2]);
end.
