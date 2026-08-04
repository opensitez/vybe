// vybe-test: pascal/pascal_inline_assembly_basm/test_asm_write_int_at_ptr
// origin: languages/pascal/tests/pascal/test_pascal_inline_assembly_basm.rs
program Test;
{$mode delphi}
uses SysUtils;
procedure WriteIntAtPtr(p: PInteger; val: Integer);
asm
  mov eax, p
  mov edx, val
  mov [eax], edx
end;
var x: Integer;
begin
  x := 0;
  WriteIntAtPtr(@x, 777);
  WriteLn(x);
end.
