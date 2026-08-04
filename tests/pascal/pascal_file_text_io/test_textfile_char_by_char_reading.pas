// vybe-test: pascal/pascal_file_text_io/test_textfile_char_by_char_reading
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; ch: Char;
begin
  AssignFile(f, 'test_chars.txt');
  Rewrite(f);
  Write(f, 'XY');
  CloseFile(f);

  Reset(f);
  Read(f, ch); WriteLn(ch);
  Read(f, ch); WriteLn(ch);
  CloseFile(f);
end.
