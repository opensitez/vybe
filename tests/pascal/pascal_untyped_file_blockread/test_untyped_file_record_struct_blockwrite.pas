// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_record_struct_blockwrite
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
type THeader = packed record
  Magic: Word;
  Version: Byte;
end;
var f: file;
    h1, h2: THeader;
    written, readBytes: Integer;
begin
  h1.Magic := $4D5A; h1.Version := 2;
  AssignFile(f, 'test_hdr.bin');
  Rewrite(f, 1);
  BlockWrite(f, h1, SizeOf(THeader), written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, h2, SizeOf(THeader), readBytes);
  WriteLn(h2.Magic);
  WriteLn(h2.Version);
  CloseFile(f);
end.
