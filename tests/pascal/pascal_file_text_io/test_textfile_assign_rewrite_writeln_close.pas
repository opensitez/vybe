// vybe-test: pascal/pascal_file_text_io/test_textfile_assign_rewrite_writeln_close
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_text_1.txt');
  Rewrite(f);
  WriteLn(f, 'Line 1');
  WriteLn(f, 'Line 2');
  CloseFile(f);
  WriteLn('Written');
end.
