// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_chunk_copy_loop
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var srcFile, dstFile: file;
    buf: array[0..63] of Byte;
    readCount, writtenCount: Integer;
begin
  buf[0] := 99; buf[63] := 88;
  AssignFile(srcFile, 'test_chunk_src.bin');
  Rewrite(srcFile, 1);
  BlockWrite(srcFile, buf[0], 64, writtenCount);
  CloseFile(srcFile);

  Reset(srcFile, 1);
  AssignFile(dstFile, 'test_chunk_dst.bin');
  Rewrite(dstFile, 1);

  repeat
    BlockRead(srcFile, buf[0], 64, readCount);
    if readCount > 0 then
      BlockWrite(dstFile, buf[0], readCount, writtenCount);
  until readCount = 0;

  CloseFile(srcFile); CloseFile(dstFile);
  WriteLn('ChunkCopyComplete');
end.
