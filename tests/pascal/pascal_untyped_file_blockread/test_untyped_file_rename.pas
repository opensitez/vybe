// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_rename
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file; b: Byte; written: Integer;
begin
  b := 77;
  AssignFile(f, 'test_orig_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);
  Rename(f, 'test_renamed_raw.bin');
  WriteLn('RenamedUntypedFile');
end.
