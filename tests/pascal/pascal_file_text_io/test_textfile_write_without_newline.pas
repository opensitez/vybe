// vybe-test: pascal/pascal_file_text_io/test_textfile_write_without_newline
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_no_nl.txt');
  Rewrite(f);
  Write(f, 'Hello ');
  Write(f, 'World');
  WriteLn(f);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
