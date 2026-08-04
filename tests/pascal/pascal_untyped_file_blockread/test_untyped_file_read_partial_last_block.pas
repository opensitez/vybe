// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_read_partial_last_block
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    buf: array[0..2] of Byte;
    readCount: Integer;
begin
  AssignFile(f, 'test_part.bin');
  Rewrite(f, 1);
  buf[0] := 1; buf[1] := 2;
  BlockWrite(f, buf[0], 2, readCount);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, buf[0], 5, readCount);
  WriteLn(readCount);
  CloseFile(f);
end.
