// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_string_bytes_blockwrite
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    text: String;
    readText: String;
    written, readBytes: Integer;
begin
  text := 'RawStringBytes';
  AssignFile(f, 'test_str_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, text[1], Length(text), written);
  CloseFile(f);

  Reset(f, 1);
  SetLength(readText, Length(text));
  BlockRead(f, readText[1], Length(text), readBytes);
  WriteLn(readText);
  CloseFile(f);
end.
