// vybe-test: pascal/pascal_typed_file_io/test_typed_file_filesize_and_filepos
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_pos.bin');
  Rewrite(f);
  val := 10; Write(f, val);
  val := 20; Write(f, val);
  val := 30; Write(f, val);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f));
  WriteLn(FilePos(f));
  Read(f, val);
  WriteLn(FilePos(f));
  CloseFile(f);
end.
