// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_erase
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; b: Byte; written: Integer;
begin
  b := 55;
  AssignFile(f, 'test_erase_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);
  Erase(f);
  WriteLn('ErasedUntypedFile');
end.
