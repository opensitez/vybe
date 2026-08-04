// vybe-test: pascal/generics_advanced/generic_buffer_write_read
// origin: languages/pascal/tests/pascal/test_generics_advanced.rs
program T;
{$mode delphi}
uses SysUtils; type TBuf<T>=class private F:T; public procedure Write(v:T); function Read:T; end; procedure TBuf<T>.Write(v:T); begin F:=v; end; function TBuf<T>.Read:T; begin Result:=F; end; var b:TBuf<Double>; begin b:=TBuf<Double>.Create; b.Write(3.5); WriteLn(b.Read>3.0); b.Free; end.
