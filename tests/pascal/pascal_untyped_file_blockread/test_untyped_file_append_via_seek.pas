// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_append_via_seek
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; b: Byte; written, readCount: Integer;
begin
  AssignFile(f, 'test_app_raw.bin');
  Rewrite(f, 1);
  b := 10; BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, FileSize(f));
  b := 20; BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
