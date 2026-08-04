// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_default_record_size_128
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
begin
  AssignFile(f, 'test_def_rec.bin');
  Rewrite(f); // Default 128 bytes
  WriteLn('DefaultRecSize128Initialized');
  CloseFile(f);
end.
