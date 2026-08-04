// vybe-test: pascal/pascal_file_text_io/test_textfile_reading_multiple_strings
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; s1, s2: String;
begin
  AssignFile(f, 'test_str2.txt');
  Rewrite(f);
  WriteLn(f, 'Alpha');
  WriteLn(f, 'Beta');
  CloseFile(f);

  Reset(f);
  ReadLn(f, s1); ReadLn(f, s2);
  WriteLn(s1 + ' & ' + s2);
  CloseFile(f);
end.
