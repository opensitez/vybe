// vybe-test: pascal/pascal_file_text_io/test_textfile_eoln_detection
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; ch: Char; count: Integer;
begin
  AssignFile(f, 'test_eoln.txt');
  Rewrite(f);
  WriteLn(f, 'ABC');
  CloseFile(f);

  Reset(f);
  count := 0;
  while not Eoln(f) do
  begin
    Read(f, ch);
    Inc(count);
  end;
  WriteLn(count);
  CloseFile(f);
end.
