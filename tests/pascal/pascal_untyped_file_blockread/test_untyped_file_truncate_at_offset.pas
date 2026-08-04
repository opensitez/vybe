// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_truncate_at_offset
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    buf: array[0..9] of Byte;
    written: Integer;
begin
  AssignFile(f, 'test_trunc_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, buf[0], 10, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 4);
  Truncate(f);
  CloseFile(f);

  Reset(f, 1);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
