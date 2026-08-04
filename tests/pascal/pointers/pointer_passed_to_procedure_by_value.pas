// vybe-test: pascal/pointers/pointer_passed_to_procedure_by_value
// origin: languages/pascal/tests/pascal/test_pointers.rs
program T;
{$mode delphi}
uses SysUtils;
procedure WriteThrough(p: ^Integer);
begin
  WriteLn(p^);
end;
var x: Integer;
begin
  x := 55;
  WriteThrough(@x);
end.
