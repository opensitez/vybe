// vybe-test: pascal/file_io/textfile_two_handles_independent_positions_same_name
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; var writer, reader: TextFile; s: string; begin Assign(writer,'core_shared.txt'); Rewrite(writer); WriteLn(writer,'shared'); Close(writer); Assign(reader,'core_shared.txt'); Reset(reader); ReadLn(reader,s); Close(reader); WriteLn(s); end.
