use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 25: Untouched Buffer Operations (FillChar, Move, ZeroMemory)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_fillchar_zeroing_integer_array() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[1..4] of Integer;
begin
  arr[1] := 10; arr[2] := 20; arr[3] := 30; arr[4] := 40;
  FillChar(arr, SizeOf(arr), 0);
  WriteLn(arr[1]);
  WriteLn(arr[4]);
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_fillchar_byte_pattern() {
    let out = run_pascal(
        r#"
program Test;
var buf: array[0..3] of Byte;
begin
  FillChar(buf[0], 4, $FF);
  WriteLn(buf[0]);
  WriteLn(buf[3]);
end.
"#,
    );
    assert_eq!(out, vec!["255", "255"]);
}

#[test]
fn test_move_copying_integer_arrays() {
    let out = run_pascal(
        r#"
program Test;
var src, dst: array[0..2] of Integer;
begin
  src[0] := 100; src[1] := 200; src[2] := 300;
  FillChar(dst, SizeOf(dst), 0);
  Move(src[0], dst[0], SizeOf(src));
  WriteLn(dst[0]);
  WriteLn(dst[1]);
  WriteLn(dst[2]);
end.
"#,
    );
    assert_eq!(out, vec!["100", "200", "300"]);
}

#[test]
fn test_move_copying_record_data() {
    let out = run_pascal(
        r#"
program Test;
type TData = record A, B: Integer; end;
var r1, r2: TData;
begin
  r1.A := 55; r1.B := 99;
  Move(r1, r2, SizeOf(TData));
  WriteLn(r2.A);
  WriteLn(r2.B);
end.
"#,
    );
    assert_eq!(out, vec!["55", "99"]);
}

#[test]
fn test_fillchar_char_fill() {
    let out = run_pascal(
        r#"
program Test;
var charBuf: array[0..4] of Char;
begin
  FillChar(charBuf[0], 4, 'X');
  charBuf[4] := #0;
  WriteLn(charBuf[0]);
  WriteLn(charBuf[3]);
end.
"#,
    );
    assert_eq!(out, vec!["X", "X"]);
}

#[test]
fn test_fillbyte_procedure() {
    let out = run_pascal(
        r#"
program Test;
var bytes: array[0..3] of Byte;
begin
  FillByte(bytes[0], 4, 128);
  WriteLn(bytes[0]);
  WriteLn(bytes[3]);
end.
"#,
    );
    assert_eq!(out, vec!["128", "128"]);
}

#[test]
fn test_fillword_procedure() {
    let out = run_pascal(
        r#"
program Test;
var words: array[0..2] of Word;
begin
  FillWord(words[0], 3, 1000);
  WriteLn(words[0]);
  WriteLn(words[2]);
end.
"#,
    );
    assert_eq!(out, vec!["1000", "1000"]);
}

#[test]
fn test_move_partial_array_slice() {
    let out = run_pascal(
        r#"
program Test;
var src, dst: array[0..4] of Integer;
begin
  src[0] := 1; src[1] := 2; src[2] := 3; src[3] := 4; src[4] := 5;
  FillChar(dst, SizeOf(dst), 0);
  Move(src[1], dst[0], SizeOf(Integer) * 3);
  WriteLn(dst[0]);
  WriteLn(dst[1]);
  WriteLn(dst[2]);
  WriteLn(dst[3]);
end.
"#,
    );
    assert_eq!(out, vec!["2", "3", "4", "0"]);
}

#[test]
fn test_untyped_buffer_parameter_processing() {
    let out = run_pascal(
        r#"
program Test;
procedure ClearBuffer(var buf; size: Integer);
begin
  FillChar(buf, size, 0);
end;
var nums: array[1..3] of Integer;
begin
  nums[1] := 10; nums[2] := 20; nums[3] := 30;
  ClearBuffer(nums, SizeOf(nums));
  WriteLn(nums[1]);
  WriteLn(nums[3]);
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_move_between_dynamic_array_and_static_array() {
    let out = run_pascal(
        r#"
program Test;
var dynArr: array of Integer;
    statArr: array[0..2] of Integer;
begin
  SetLength(dynArr, 3);
  dynArr[0] := 7; dynArr[1] := 14; dynArr[2] := 21;
  Move(dynArr[0], statArr[0], SizeOf(Integer) * 3);
  WriteLn(statArr[0]);
  WriteLn(statArr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["7", "21"]);
}

#[test]
fn test_move_overlapping_buffer_shift() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[0..4] of Integer;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30; arr[3] := 40; arr[4] := 50;
  Move(arr[0], arr[1], SizeOf(Integer) * 3);
  WriteLn(arr[1]);
  WriteLn(arr[2]);
  WriteLn(arr[3]);
end.
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_fillchar_clearing_record_before_usage() {
    let out = run_pascal(
        r#"
program Test;
type THeader = record
  Magic: Integer;
  Version: Integer;
end;
var h: THeader;
begin
  FillChar(h, SizeOf(THeader), 0);
  WriteLn(h.Magic);
  WriteLn(h.Version);
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_move_string_characters_to_byte_array() {
    let out = run_pascal(
        r#"
program Test;
var strVal: String;
    bytes: array[0..4] of Byte;
begin
  strVal := 'ABCDE';
  Move(strVal[1], bytes[0], 5);
  WriteLn(bytes[0]);
  WriteLn(bytes[1]);
  WriteLn(bytes[2]);
end.
"#,
    );
    assert_eq!(out, vec!["65", "66", "67"]);
}

#[test]
fn test_untyped_move_wrapper_procedure() {
    let out = run_pascal(
        r#"
program Test;
procedure CopyData(const src; var dst; count: Integer);
begin
  Move(src, dst, count);
end;
var a, b: Integer;
begin
  a := 999;
  CopyData(a, b, SizeOf(Integer));
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_fillchar_multidimensional_matrix() {
    let out = run_pascal(
        r#"
program Test;
var mat: array[0..1, 0..1] of Integer;
begin
  mat[0, 0] := 1; mat[0, 1] := 2;
  mat[1, 0] := 3; mat[1, 1] := 4;
  FillChar(mat, SizeOf(mat), 0);
  WriteLn(mat[0, 0]);
  WriteLn(mat[1, 1]);
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_raw_byte_compare_after_move() {
    let out = run_pascal(
        r#"
program Test;
function SameBytes(const a, b; len: Integer): Boolean;
var pa, pb: PByte; i: Integer;
begin
  pa := @a; pb := @b; Result := True;
  for i := 0 to len - 1 do
  begin
    if pa^ <> pb^ then begin Result := False; Break; end;
    Inc(pa); Inc(pb);
  end;
end;
var x, y: Integer;
begin
  x := 12345;
  Move(x, y, SizeOf(Integer));
  WriteLn(SameBytes(x, y, SizeOf(Integer)));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_move_with_pbyte_offset_source() {
    let out = run_pascal(
        r#"
program Test;
var buf: array[0..5] of Byte;
    outVal: Integer;
begin
  buf[0] := 0; buf[1] := 0; buf[2] := 42; buf[3] := 0; buf[4] := 0; buf[5] := 0;
  Move(buf[2], outVal, SizeOf(Integer));
  WriteLn(outVal);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_fillword_filling_shortint_array() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[0..3] of Word;
begin
  FillWord(arr[0], 4, 12345);
  WriteLn(arr[0]);
  WriteLn(arr[3]);
end.
"#,
    );
    assert_eq!(out, vec!["12345", "12345"]);
}

#[test]
fn test_move_copying_real_array() {
    let out = run_pascal(
        r#"
program Test;
var rSrc, rDst: array[0..1] of Real;
begin
  rSrc[0] := 1.5; rSrc[1] := 2.5;
  Move(rSrc[0], rDst[0], SizeOf(rSrc));
  WriteLn(rDst[0]);
  WriteLn(rDst[1]);
end.
"#,
    );
    assert_eq!(out, vec!["1.5", "2.5"]);
}

#[test]
fn test_fillchar_boolean_array() {
    let out = run_pascal(
        r#"
program Test;
var flags: array[0..2] of Boolean;
begin
  flags[0] := True; flags[1] := True; flags[2] := True;
  FillChar(flags, SizeOf(flags), 0);
  WriteLn(flags[0]);
  WriteLn(flags[2]);
end.
"#,
    );
    assert_eq!(out, vec!["False", "False"]);
}
