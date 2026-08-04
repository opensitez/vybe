// vybe-test: pascal/pascal_file_text_io/test_textfile_empty_lines_reading
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; l1, l2: String;
begin
  AssignFile(f, 'test_empty.txt');
  Rewrite(f);
  WriteLn(f, '');
  WriteLn(f, 'SecondLine');
  CloseFile(f);

  Reset(f);
  ReadLn(f, l1);
  ReadLn(f, l2);
  WriteLn(Length(l1));
  WriteLn(l2);
  CloseFile(f);
end.
