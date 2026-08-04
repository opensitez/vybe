// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_eof_check
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; b: Byte; written, readCount: Integer;
begin
  b := 1;
  AssignFile(f, 'test_eof_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, b, 1, readCount);
  WriteLn(Eof(f));
  CloseFile(f);
end.
