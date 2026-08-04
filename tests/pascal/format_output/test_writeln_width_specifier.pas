// vybe-test: pascal/format_output/test_writeln_width_specifier
// origin: languages/pascal/tests/pascal/test_format_output.rs
program T;
{$mode delphi}
uses SysUtils;
var
  n: Integer;
  f: Double;
begin
  n := 42;
  WriteLn(n:8);
  f := 3.14;
  WriteLn(f:8:2);
end.
