// vybe-test: pascal/file_io/textfile_append_preserves_existing_content
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'core_append.txt'); Rewrite(f); WriteLn(f,'first'); Close(f); Append(f); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end.
