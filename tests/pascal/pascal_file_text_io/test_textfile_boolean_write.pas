// vybe-test: pascal/pascal_file_text_io/test_textfile_boolean_write
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_bool.txt');
  Rewrite(f);
  WriteLn(f, True);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
