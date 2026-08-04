// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_seek_block
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    buf: array[0..3] of Byte;
    val: Byte; written, readCount: Integer;
begin
  buf[0] := 1; buf[1] := 2; buf[2] := 3; buf[3] := 4;
  AssignFile(f, 'test_seek_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, buf[0], 4, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 2);
  BlockRead(f, val, 1, readCount);
  WriteLn(val);
  CloseFile(f);
end.
