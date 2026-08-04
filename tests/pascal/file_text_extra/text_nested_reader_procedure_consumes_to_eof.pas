// vybe-test: pascal/file_text_extra/text_nested_reader_procedure_consumes_to_eof
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; procedure ReadAll(var total: Integer); var s: string; begin total := 0; while not Eof(f) do begin ReadLn(f,s); total := total + Length(s); end; end; var n: Integer; begin Assign(f,'text_nested.txt'); Rewrite(f); WriteLn(f,'ab'); WriteLn(f,'cde'); Close(f); Reset(f); ReadAll(n); Close(f); WriteLn(n); end.
