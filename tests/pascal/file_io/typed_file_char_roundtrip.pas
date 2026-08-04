// vybe-test: pascal/file_io/typed_file_char_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io.rs
program T;
{$mode delphi}
uses SysUtils; type TCharFile = file of Char; var f: TCharFile; c: Char; begin Assign(f,'core_chars.dat'); Rewrite(f); c := 'Z'; Write(f,c); Close(f); Reset(f); c := 'X'; Read(f,c); Close(f); WriteLn(c); end.
