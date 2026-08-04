// vybe-test: pascal/pascal_file_text_io/test_textfile_procedure_parameter
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
procedure WriteHeader(var f: TextFile);
begin
  WriteLn(f, '=== HEADER ===');
end;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_proc.txt');
  Rewrite(f);
  WriteHeader(f);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
