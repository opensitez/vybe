// vybe-test: pascal/file_text_extra/text_filter_loop_reads_each_line_once
// origin: languages/pascal/tests/pascal/test_file_text_extra.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; n: Integer; begin Assign(f,'text_filter.txt'); Rewrite(f); WriteLn(f,'ERR one'); WriteLn(f,'OK two'); WriteLn(f,'ERR three'); Close(f); Reset(f); n := 0; while not Eof(f) do begin ReadLn(f,s); if Copy(s,1,3) = 'ERR' then n := n + 1; end; Close(f); WriteLn(n); end.
