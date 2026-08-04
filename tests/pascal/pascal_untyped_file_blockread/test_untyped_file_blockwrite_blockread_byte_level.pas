// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_blockwrite_blockread_byte_level
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    writeBuf, readBuf: array[0..3] of Byte;
    written, readBytes: Integer;
begin
  writeBuf[0] := 10; writeBuf[1] := 20; writeBuf[2] := 30; writeBuf[3] := 40;
  AssignFile(f, 'test_untyped_1.bin');
  Rewrite(f, 1);
  BlockWrite(f, writeBuf[0], 4, written);
  CloseFile(f);
  WriteLn(written);

  Reset(f, 1);
  BlockRead(f, readBuf[0], 4, readBytes);
  CloseFile(f);
  WriteLn(readBytes);
  WriteLn(readBuf[0]);
  WriteLn(readBuf[3]);
end.
