use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 89: Inline Assembly (BASM / asm...end blocks)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_asm_basic_add_registers() {
    let out = run_pascal(r#"
program Test;
function AddAsm(a, b: Integer): Integer;
asm
  mov eax, a
  add eax, b
  mov Result, eax
end;
begin
  WriteLn(AddAsm(15, 25));
end.
"#);
    assert_eq!(out, vec!["40"]);
}

#[test]
fn test_asm_subtraction() {
    let out = run_pascal(r#"
program Test;
function SubAsm(a, b: Integer): Integer;
asm
  mov eax, a
  sub eax, b
  mov Result, eax
end;
begin
  WriteLn(SubAsm(100, 30));
end.
"#);
    assert_eq!(out, vec!["70"]);
}

#[test]
fn test_asm_multiplication_imul() {
    let out = run_pascal(r#"
program Test;
function MulAsm(a, b: Integer): Integer;
asm
  mov eax, a
  imul eax, b
  mov Result, eax
end;
begin
  WriteLn(MulAsm(6, 7));
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_asm_bitwise_and_or_xor() {
    let out = run_pascal(r#"
program Test;
function BitOpsAsm(a, b: Integer): Integer;
asm
  mov eax, a
  and eax, b
  mov Result, eax
end;
begin
  WriteLn(BitOpsAsm($0F, $33));
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_asm_increment_decrement() {
    let out = run_pascal(r#"
program Test;
function IncDecAsm(val: Integer): Integer;
asm
  mov eax, val
  inc eax
  inc eax
  dec eax
  mov Result, eax
end;
begin
  WriteLn(IncDecAsm(10));
end.
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_asm_shift_left_right() {
    let out = run_pascal(r#"
program Test;
function ShiftLeftAsm(val, count: Integer): Integer;
asm
  mov eax, val
  mov ecx, count
  shl eax, cl
  mov Result, eax
end;
begin
  WriteLn(ShiftLeftAsm(8, 2));
end.
"#);
    assert_eq!(out, vec!["32"]);
}

#[test]
fn test_asm_negate_register() {
    let out = run_pascal(r#"
program Test;
function NegAsm(val: Integer): Integer;
asm
  mov eax, val
  neg eax
  mov Result, eax
end;
begin
  WriteLn(NegAsm(50));
end.
"#);
    assert_eq!(out, vec!["-50"]);
}

#[test]
fn test_asm_bitwise_not() {
    let out = run_pascal(r#"
program Test;
function NotAsm(val: Byte): Byte;
asm
  mov al, val
  not al
  mov Result, al
end;
begin
  WriteLn(NotAsm($00));
end.
"#);
    assert_eq!(out, vec!["255"]);
}

#[test]
fn test_asm_loop_with_labels() {
    let out = run_pascal(r#"
program Test;
function SumAsm(n: Integer): Integer;
asm
  mov ecx, n
  xor eax, eax
@@Loop:
  add eax, ecx
  dec ecx
  jnz @@Loop
  mov Result, eax
end;
begin
  WriteLn(SumAsm(5)); // 5+4+3+2+1 = 15
end.
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_asm_pointer_dereference() {
    let out = run_pascal(r#"
program Test;
function ReadIntAtPtr(p: PInteger): Integer;
asm
  mov eax, p
  mov eax, [eax]
  mov Result, eax
end;
var x: Integer;
begin
  x := 999;
  WriteLn(ReadIntAtPtr(@x));
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_asm_write_int_at_ptr() {
    let out = run_pascal(r#"
program Test;
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
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_asm_push_pop_registers() {
    let out = run_pascal(r#"
program Test;
function PushPopAsm(a, b: Integer): Integer;
asm
  push ebx
  mov eax, a
  mov ebx, b
  add eax, ebx
  pop ebx
  mov Result, eax
end;
begin
  WriteLn(PushPopAsm(40, 60));
end.
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_asm_conditional_move_cmovz() {
    let out = run_pascal(r#"
program Test;
function MaxAsm(a, b: Integer): Integer;
asm
  mov eax, a
  cmp eax, b
  jge @@Done
  mov eax, b
@@Done:
  mov Result, eax
end;
begin
  WriteLn(MaxAsm(10, 25));
  WriteLn(MaxAsm(50, 20));
end.
"#);
    assert_eq!(out, vec!["25", "50"]);
}

#[test]
fn test_asm_xor_self_clear() {
    let out = run_pascal(r#"
program Test;
function ZeroAsm: Integer;
asm
  xor eax, eax
  mov Result, eax
end;
begin
  WriteLn(ZeroAsm);
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_asm_bsf_bit_scan_forward() {
    let out = run_pascal(r#"
program Test;
function FirstBitSetAsm(val: Integer): Integer;
asm
  mov eax, val
  bsf eax, eax
  mov Result, eax
end;
begin
  WriteLn(FirstBitSetAsm(16)); // Bit 4 set (16 = 2^4)
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_asm_bsr_bit_scan_reverse() {
    let out = run_pascal(r#"
program Test;
function LastBitSetAsm(val: Integer): Integer;
asm
  mov eax, val
  bsr eax, eax
  mov Result, eax
end;
begin
  WriteLn(LastBitSetAsm(33)); // Bit 5 set (32 = 2^5)
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_asm_swap_two_vars() {
    let out = run_pascal(r#"
program Test;
procedure SwapAsm(var a, b: Integer);
asm
  mov eax, a
  mov edx, b
  mov ecx, [eax]
  mov esi, [edx]
  mov [eax], esi
  mov [edx], ecx
end;
var x, y: Integer;
begin
  x := 10; y := 20;
  SwapAsm(x, y);
  WriteLn(x.ToString + ',' + y.ToString);
end.
"#);
    assert_eq!(out, vec!["20,10"]);
}

#[test]
fn test_asm_nop_instruction() {
    let out = run_pascal(r#"
program Test;
procedure NopProc;
asm
  nop
  nop
end;
begin
  NopProc;
  WriteLn('NopExecuted');
end.
"#);
    assert_eq!(out, vec!["NopExecuted"]);
}

#[test]
fn test_asm_byte_array_sum() {
    let out = run_pascal(r#"
program Test;
function SumBytesAsm(p: PByte; count: Integer): Integer;
asm
  mov esi, p
  mov ecx, count
  xor eax, eax
  test ecx, ecx
  jz @@Done
@@Loop:
  movzx edx, byte ptr [esi]
  add eax, edx
  inc esi
  dec ecx
  jnz @@Loop
@@Done:
  mov Result, eax
end;
var arr: array[0..2] of Byte;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  WriteLn(SumBytesAsm(@arr[0], 3));
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_asm_fast_square() {
    let out = run_pascal(r#"
program Test;
function SquareAsm(val: Integer): Integer;
asm
  mov eax, val
  imul eax, eax
  mov Result, eax
end;
begin
  WriteLn(SquareAsm(9));
end.
"#);
    assert_eq!(out, vec!["81"]);
}
