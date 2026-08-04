// vybe-test: pascal/pascal_file_text_io/test_textfile_flush_buffer
// origin: languages/pascal/tests/pascal/test_pascal_file_text_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_flush.txt');
  Rewrite(f);
  WriteLn(f, 'FlushedData');
  Flush(f);
  CloseFile(f);
  WriteLn('FlushedSuccessfully');
end.
