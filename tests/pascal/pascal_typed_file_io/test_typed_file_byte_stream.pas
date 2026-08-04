// vybe-test: pascal/pascal_typed_file_io/test_typed_file_byte_stream
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Byte; b: Byte;
begin
  AssignFile(f, 'test_byte.bin');
  Rewrite(f);
  b := $AB; Write(f, b);
  b := $CD; Write(f, b);
  CloseFile(f);

  Reset(f);
  Read(f, b); WriteLn(HexStr(b, 2));
  Read(f, b); WriteLn(HexStr(b, 2));
  CloseFile(f);
end.
