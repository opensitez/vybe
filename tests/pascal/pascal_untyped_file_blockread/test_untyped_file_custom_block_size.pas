// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_custom_block_size
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBlock = array[0..15] of Byte;
var f: file;
    blk: TBlock;
    written, readBlocks: Integer;
begin
  FillChar(blk, SizeOf(TBlock), 65);
  AssignFile(f, 'test_block_16.bin');
  Rewrite(f, 16);
  BlockWrite(f, blk, 1, written);
  CloseFile(f);

  Reset(f, 16);
  BlockRead(f, blk, 1, readBlocks);
  CloseFile(f);
  WriteLn(readBlocks);
  WriteLn(blk[0]);
end.
