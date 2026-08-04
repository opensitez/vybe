// vybe-test: pascal/pascal_typed_file_io/test_typed_file_seek_random_access
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_seek.bin');
  Rewrite(f);
  val := 10; Write(f, val);
  val := 20; Write(f, val);
  val := 30; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
