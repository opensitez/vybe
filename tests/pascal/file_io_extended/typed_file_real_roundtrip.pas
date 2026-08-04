// vybe-test: pascal/file_io_extended/typed_file_real_roundtrip
// origin: languages/pascal/tests/pascal/test_file_io_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TRealFile = file of Real; var f: TRealFile; r: Real; begin Assign(f,'ext_real.dat'); Rewrite(f); r := 2.5; Write(f,r); Close(f); Reset(f); r := 0; Read(f,r); Close(f); WriteLn(Trunc(r*10)); end.
