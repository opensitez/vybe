// vybe-test: pascal/pascal_typed_file_io/test_typed_file_eof_iteration
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val, sum: Integer;
begin
  AssignFile(f, 'test_sum.bin');
  Rewrite(f);
  val := 5; Write(f, val);
  val := 15; Write(f, val);
  val := 25; Write(f, val);
  CloseFile(f);

  Reset(f);
  sum := 0;
  while not Eof(f) do
  begin
    Read(f, val);
    sum := sum + val;
  end;
  WriteLn(sum);
  CloseFile(f);
end.
