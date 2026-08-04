// vybe-test: pascal/pascal_file_text_io/test_textfile_read_formatted_values
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; a, b: Integer;
begin
  AssignFile(f, 'test_fmt.txt');
  Rewrite(f);
  WriteLn(f, '10 20');
  CloseFile(f);

  Reset(f);
  Read(f, a); Read(f, b);
  WriteLn(a + b);
  CloseFile(f);
end.
