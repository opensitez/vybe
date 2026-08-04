// vybe-test: pascal/pascal_typed_file_io/test_typed_file_word_stream
// origin: languages/pascal/tests/pascal/test_pascal_typed_file_io.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file of Word; w: Word;
begin
  AssignFile(f, 'test_word.bin');
  Rewrite(f);
  w := 65000; Write(f, w);
  CloseFile(f);

  Reset(f);
  Read(f, w); WriteLn(w);
  CloseFile(f);
end.
