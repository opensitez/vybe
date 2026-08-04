// vybe-test: pascal/pascal_typed_file_io/test_typed_file_truncate_at_position
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_trunc.bin');
  Rewrite(f);
  val := 1; Write(f, val);
  val := 2; Write(f, val);
  val := 3; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Truncate(f);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
