// vybe-test: pascal/file_io/textfile_function_reads_global_handle
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; function Load: string; var s: string; begin Reset(f); ReadLn(f,s); Close(f); Result := s; end; begin Assign(f,'core_func.txt'); Rewrite(f); WriteLn(f,'from func'); Close(f); WriteLn(Load); end.
