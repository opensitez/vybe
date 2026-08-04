// vybe-test: pascal/file_text_extra/text_eof_while_loop_counts_blank_and_nonblank_lines
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; n: Integer; begin Assign(f,'text_count.txt'); Rewrite(f); WriteLn(f); WriteLn(f,'a'); WriteLn(f,'bb'); Close(f); Reset(f); n := 0; while not Eof(f) do begin ReadLn(f,s); n := n + 1; end; Close(f); WriteLn(n); end.
