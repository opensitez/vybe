// vybe-test: pascal/file_io_extended/flush_preserves_data_before_close
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; var f: TextFile; s: string; begin Assign(f,'ext_flush.txt'); Rewrite(f); WriteLn(f,'flushed'); Flush(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end.
