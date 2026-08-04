// vybe-test: pascal/pascal_typed_file_io/test_typed_file_integer_write_read
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_int.bin');
  Rewrite(f);
  val := 100; Write(f, val);
  val := 200; Write(f, val);
  CloseFile(f);

  Reset(f);
  Read(f, val); WriteLn(val);
  Read(f, val); WriteLn(val);
  CloseFile(f);
end.
