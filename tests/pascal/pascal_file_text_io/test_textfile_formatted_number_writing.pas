// vybe-test: pascal/pascal_file_text_io/test_textfile_formatted_number_writing
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile; val: Integer; r: Real;
begin
  AssignFile(f, 'test_num.txt');
  Rewrite(f);
  val := 42; r := 3.14;
  WriteLn(f, val);
  WriteLn(f, r:0:2);
  CloseFile(f);

  Reset(f);
  ReadLn(f, val);
  WriteLn(val);
  CloseFile(f);
end.
