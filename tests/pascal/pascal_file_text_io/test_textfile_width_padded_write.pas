// vybe-test: pascal/pascal_file_text_io/test_textfile_width_padded_write
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_pad.txt');
  Rewrite(f);
  WriteLn(f, 99:5);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn('[' + line + ']');
  CloseFile(f);
end.
