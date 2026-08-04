// vybe-test: pascal/format_output/test_write_specifier_aligned
// origin: languages/pascal/tests/pascal/test_format_output.rs
program T;
{$mode delphi}
uses SysUtils;
begin
  Write(1:4);
  Write(2:4);
  Write(3:4);
  WriteLn('');
end.
