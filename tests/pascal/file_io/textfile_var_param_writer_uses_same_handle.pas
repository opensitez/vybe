// vybe-test: pascal/file_io/textfile_var_param_writer_uses_same_handle
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; procedure Save(var f: TextFile; msg: string); begin WriteLn(f,msg); end; var f: TextFile; s: string; begin Assign(f,'core_proc.txt'); Rewrite(f); Save(f,'via proc'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
