use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 23: Pointer Arithmetic & Offset Calculations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_pbyte_inc_dec() {
    let out = run_pascal(r#"
program Test;
var bytes: array[0..2] of Byte;
    pb: PByte;
begin
  bytes[0] := 10; bytes[1] := 20; bytes[2] := 30;
  pb := @bytes[0];
  WriteLn(pb^);
  Inc(pb);
  WriteLn(pb^);
  Inc(pb);
  WriteLn(pb^);
end.
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_pinteger_pointer_arithmetic_stride() {
    let out = run_pascal(r#"
program Test;
var ints: array[0..2] of Integer;
    pi: PInteger;
begin
  ints[0] := 100; ints[1] := 200; ints[2] := 300;
  pi := @ints[0];
  Inc(pi);
  WriteLn(pi^);
  Inc(pi, 1);
  WriteLn(pi^);
end.
"#);
    assert_eq!(out, vec!["200", "300"]);
}

#[test]
fn test_pointer_array_indexing_syntax() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p: PInteger;
begin
  arr[0] := 5; arr[1] := 15; arr[2] := 25;
  p := @arr[0];
  WriteLn(p[0]);
  WriteLn(p[1]);
  WriteLn(p[2]);
end.
"#);
    assert_eq!(out, vec!["5", "15", "25"]);
}

#[test]
fn test_pointer_subtraction_byte_distance() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..4] of Byte;
    p1, p2: PByte;
    diff: NativeInt;
begin
  p1 := @arr[0];
  p2 := @arr[4];
  diff := NativeInt(p2) - NativeInt(p1);
  WriteLn(diff);
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_pchar_pointer_addition() {
    let out = run_pascal(r#"
program Test;
var strBuf: array[0..5] of Char;
    pc: PChar;
begin
  strBuf[0] := 'H'; strBuf[1] := 'i'; strBuf[2] := #0;
  pc := @strBuf[0];
  Inc(pc);
  WriteLn(pc^);
end.
"#);
    assert_eq!(out, vec!["i"]);
}

#[test]
fn test_pointer_dec_loop_iteration() {
    let out = run_pascal(r#"
program Test;
var nums: array[0..2] of Integer;
    p: PInteger;
begin
  nums[0] := 1; nums[1] := 2; nums[2] := 3;
  p := @nums[2];
  WriteLn(p^);
  Dec(p);
  WriteLn(p^);
  Dec(p);
  WriteLn(p^);
end.
"#);
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn test_pointer_offset_with_pbyte_casting() {
    let out = run_pascal(r#"
program Test;
type TData = record
  A: Integer;
  B: Integer;
end;
var rec: TData;
    pBase: Pointer;
    pB: PInteger;
begin
  rec.A := 111; rec.B := 222;
  pBase := @rec;
  pB := PInteger(PByte(pBase) + SizeOf(Integer));
  WriteLn(pB^);
end.
"#);
    assert_eq!(out, vec!["222"]);
}

#[test]
fn test_pointer_comparison_boundary_check() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p, pEnd: PInteger;
    sum: Integer;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  p := @arr[0];
  pEnd := @arr[2];
  sum := 0;
  while NativeInt(p) <= NativeInt(pEnd) do
  begin
    sum := sum + p^;
    Inc(p);
  end;
  WriteLn(sum);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_pointer_pred_and_succ() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p: PInteger;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  p := @arr[1];
  WriteLn(Pred(p)^);
  WriteLn(Succ(p)^);
end.
"#);
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn test_pointer_arithmetic_custom_step() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..4] of Integer;
    p: PInteger;
begin
  arr[0] := 100; arr[2] := 300; arr[4] := 500;
  p := @arr[0];
  Inc(p, 2);
  WriteLn(p^);
  Inc(p, 2);
  WriteLn(p^);
end.
"#);
    assert_eq!(out, vec!["300", "500"]);
}

#[test]
fn test_pointer_arithmetic_record_array() {
    let out = run_pascal(r#"
program Test;
type TPoint = record X, Y: Integer; end;
type PPoint = ^TPoint;
var pts: array[0..1] of TPoint;
    p: PPoint;
begin
  pts[0].X := 1; pts[0].Y := 2;
  pts[1].X := 3; pts[1].Y := 4;
  p := @pts[0];
  Inc(p);
  WriteLn(p^.X);
  WriteLn(p^.Y);
end.
"#);
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn test_pointer_addition_operator_syntax() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p1, p2: PInteger;
begin
  arr[0] := 11; arr[1] := 22; arr[2] := 33;
  p1 := @arr[0];
  p2 := p1 + 2;
  WriteLn(p2^);
end.
"#);
    assert_eq!(out, vec!["33"]);
}

#[test]
fn test_pointer_subtraction_operator_syntax() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p1, p2: PInteger;
begin
  arr[0] := 11; arr[1] := 22; arr[2] := 33;
  p1 := @arr[2];
  p2 := p1 - 2;
  WriteLn(p2^);
end.
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_pointer_arithmetic_2d_matrix_offset() {
    let out = run_pascal(r#"
program Test;
var matrix: array[0..1, 0..1] of Integer;
    pBase: PInteger;
    val: Integer;
begin
  matrix[0, 0] := 1; matrix[0, 1] := 2;
  matrix[1, 0] := 3; matrix[1, 1] := 4;
  pBase := @matrix[0, 0];
  val := (pBase + (1 * 2 + 1))^;
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_pointer_arithmetic_byte_buffer_accumulation() {
    let out = run_pascal(r#"
program Test;
var buf: array[0..3] of Byte;
    pb: PByte;
    sum, i: Integer;
begin
  buf[0] := 5; buf[1] := 10; buf[2] := 15; buf[3] := 20;
  pb := @buf[0];
  sum := 0;
  for i := 0 to 3 do
  begin
    sum := sum + pb^;
    Inc(pb);
  end;
  WriteLn(sum);
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_pointer_arithmetic_real_stride() {
    let out = run_pascal(r#"
program Test;
var floats: array[0..1] of Real;
    pr: PReal;
begin
  floats[0] := 1.5; floats[1] := 3.5;
  pr := @floats[0];
  Inc(pr);
  WriteLn(pr^);
end.
"#);
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn test_pointer_indexing_mutation() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p: PInteger;
begin
  p := @arr[0];
  p[0] := 100;
  p[1] := 200;
  p[2] := 300;
  WriteLn(arr[0] + arr[1] + arr[2]);
end.
"#);
    assert_eq!(out, vec!["600"]);
}

#[test]
fn test_pointer_distance_element_count() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..9] of Integer;
    pStart, pEnd: PInteger;
    count: NativeInt;
begin
  pStart := @arr[0];
  pEnd := @arr[5];
  count := pEnd - pStart;
  WriteLn(count);
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_pointer_offset_with_negative_index() {
    let out = run_pascal(r#"
program Test;
var arr: array[0..2] of Integer;
    p: PInteger;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  p := @arr[2];
  WriteLn(p[-1]);
  WriteLn(p[-2]);
end.
"#);
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn test_pointer_arithmetic_boolean_array() {
    let out = run_pascal(r#"
program Test;
var flags: array[0..2] of Boolean;
    pb: PBoolean;
begin
  flags[0] := False; flags[1] := True; flags[2] := False;
  pb := @flags[0];
  Inc(pb);
  WriteLn(pb^);
end.
"#);
    assert_eq!(out, vec!["True"]);
}
