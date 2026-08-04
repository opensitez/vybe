// vybe-test: pascal/file_io_extended/typed_file_byte_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TByteFile = file of Byte; var f: TByteFile; b: Byte; begin Assign(f,'ext_byte.dat'); Rewrite(f); b := 255; Write(f,b); Close(f); Reset(f); b := 0; Read(f,b); Close(f); WriteLn(b); end.
