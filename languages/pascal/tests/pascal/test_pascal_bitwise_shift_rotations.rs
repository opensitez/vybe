use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 97: Bitwise Operations, Shifts & Rotations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_bit_shl_shr_basic() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer;
begin
  val := 1 shl 4; // 16
  WriteLn(val);
  val := val shr 2; // 4
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["16", "4"]);
}

#[test]
fn test_bit_and_or_xor_not() {
    let out = run_pascal(
        r#"
program Test;
var a, b: Byte;
begin
  a := $0F; b := $F0;
  WriteLn(HexStr(a and b, 2));
  WriteLn(HexStr(a or b, 2));
  WriteLn(HexStr(a xor b, 2));
  WriteLn(HexStr(not a, 2));
end.
"#,
    );
    assert_eq!(out, vec!["00", "FF", "FF", "F0"]);
}

#[test]
fn test_bit_lo_hi_bytes() {
    let out = run_pascal(
        r#"
program Test;
var w: Word;
begin
  w := $1234;
  WriteLn(HexStr(Lo(w), 2));
  WriteLn(HexStr(Hi(w), 2));
end.
"#,
    );
    assert_eq!(out, vec!["34", "12"]);
}

#[test]
fn test_bit_swap_bytes() {
    let out = run_pascal(
        r#"
program Test;
var w: Word;
begin
  w := $1234;
  w := Swap(w);
  WriteLn(HexStr(w, 4));
end.
"#,
    );
    assert_eq!(out, vec!["3412"]);
}

#[test]
fn test_bit_rolbyte_rorbyte() {
    let out = run_pascal(
        r#"
program Test;
uses System;
function RolByte(val, count: Byte): Byte;
begin
  Result := (val shl count) or (val shr (8 - count));
end;
function RorByte(val, count: Byte): Byte;
begin
  Result := (val shr count) or (val shl (8 - count));
end;
begin
  WriteLn(HexStr(RolByte($80, 1), 2)); // $80 rol 1 -> $01
  WriteLn(HexStr(RorByte($01, 1), 2)); // $01 ror 1 -> $80
end.
"#,
    );
    assert_eq!(out, vec!["01", "80"]);
}

#[test]
fn test_bit_rolword_rorword() {
    let out = run_pascal(
        r#"
program Test;
function RolWord(val: Word; count: Byte): Word;
begin
  Result := (val shl count) or (val shr (16 - count));
end;
begin
  WriteLn(HexStr(RolWord($8000, 1), 4)); // $8000 rol 1 -> $0001
end.
"#,
    );
    assert_eq!(out, vec!["0001"]);
}

#[test]
fn test_bit_setbit_clearbit_togglebit() {
    let out = run_pascal(
        r#"
program Test;
function SetBit(val, bitIdx: Integer): Integer; begin Result := val or (1 shl bitIdx); end;
function ClearBit(val, bitIdx: Integer): Integer; begin Result := val and not (1 shl bitIdx); end;
function ToggleBit(val, bitIdx: Integer): Integer; begin Result := val xor (1 shl bitIdx); end;

var x: Integer;
begin
  x := 0;
  x := SetBit(x, 3); // 8
  WriteLn(x);
  x := ToggleBit(x, 3); // 0
  WriteLn(x);
  x := SetBit(x, 1); // 2
  x := ClearBit(x, 1); // 0
  WriteLn(x);
end.
"#,
    );
    assert_eq!(out, vec!["8", "0", "0"]);
}

#[test]
fn test_bit_testbit_check() {
    let out = run_pascal(
        r#"
program Test;
function TestBit(val, bitIdx: Integer): Boolean;
begin
  Result := (val and (1 shl bitIdx)) <> 0;
end;
begin
  WriteLn(TestBit(8, 3));
  WriteLn(TestBit(8, 2));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_bit_mask_generation() {
    let out = run_pascal(
        r#"
program Test;
function MakeMask(bits: Integer): Integer;
begin
  Result := (1 shl bits) - 1;
end;
begin
  WriteLn(MakeMask(4)); // 15 ($0F)
  WriteLn(MakeMask(8)); // 255 ($FF)
end.
"#,
    );
    assert_eq!(out, vec!["15", "255"]);
}

#[test]
fn test_bit_loword_hiword() {
    let out = run_pascal(
        r#"
program Test;
var card: Cardinal;
begin
  card := $12345678;
  WriteLn(HexStr(LoWord(card), 4));
  WriteLn(HexStr(HiWord(card), 4));
end.
"#,
    );
    assert_eq!(out, vec!["5678", "1234"]);
}

#[test]
fn test_bit_intto_binary_string() {
    let out = run_pascal(
        r#"
program Test;
function IntToBin(val, bits: Integer): String;
var i: Integer;
begin
  Result := '';
  for i := bits - 1 downto 0 do
    if (val and (1 shl i)) <> 0 then Result := Result + '1'
    else Result := Result + '0';
end;
begin
  WriteLn(IntToBin(10, 8)); // 10 = 00001010
end.
"#,
    );
    assert_eq!(out, vec!["00001010"]);
}

#[test]
fn test_bit_popcount_ones_count() {
    let out = run_pascal(
        r#"
program Test;
function PopCount(val: Cardinal): Integer;
begin
  Result := 0;
  while val > 0 do
  begin
    Inc(Result, val and 1);
    val := val shr 1;
  end;
end;
begin
  WriteLn(PopCount($0F)); // 4
  WriteLn(PopCount($A5)); // 10100101 -> 4
end.
"#,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn test_bit_shl_multiplication_by_power_of_two() {
    let out = run_pascal(
        r#"
program Test;
var x: Integer;
begin
  x := 5;
  WriteLn(x shl 3); // 5 * 8 = 40
end.
"#,
    );
    assert_eq!(out, vec!["40"]);
}

#[test]
fn test_bit_shr_division_by_power_of_two() {
    let out = run_pascal(
        r#"
program Test;
var x: Integer;
begin
  x := 64;
  WriteLn(x shr 4); // 64 / 16 = 4
end.
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_bit_int64_shift() {
    let out = run_pascal(
        r#"
program Test;
var v64: Int64;
begin
  v64 := Int64(1) shl 40;
  WriteLn(v64 > 1000000000000);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_bit_byte_swap_int32() {
    let out = run_pascal(
        r#"
program Test;
function Swap32(val: Cardinal): Cardinal;
begin
  Result := ((val and $000000FF) shl 24) or
            ((val and $0000FF00) shl 8)  or
            ((val and $00FF0000) shr 8)  or
            ((val and $FF000000) shr 24);
end;
begin
  WriteLn(HexStr(Swap32($12345678), 8));
end.
"#,
    );
    assert_eq!(out, vec!["78563412"]);
}

#[test]
fn test_bit_sign_extension_check() {
    let out = run_pascal(
        r#"
program Test;
var sByte: ShortInt; i: Integer;
begin
  sByte := -1; // $FF
  i := Integer(sByte);
  WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn test_bit_mask_extract_field() {
    let out = run_pascal(
        r#"
program Test;
function ExtractBits(val, startBit, numBits: Integer): Integer;
var mask: Integer;
begin
  mask := (1 shl numBits) - 1;
  Result := (val shr startBit) and mask;
end;
begin
  // Extract bits 4..7 from $AB ($A = 10, $B = 11)
  WriteLn(ExtractBits($AB, 4, 4));
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_bit_is_power_of_two() {
    let out = run_pascal(
        r#"
program Test;
function IsPowerOfTwo(val: Cardinal): Boolean;
begin
  Result := (val > 0) and ((val and (val - 1)) = 0);
end;
begin
  WriteLn(IsPowerOfTwo(16));
  WriteLn(IsPowerOfTwo(18));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_bit_parity_check() {
    let out = run_pascal(
        r#"
program Test;
function IsEvenParity(val: Byte): Boolean;
var count, i: Integer;
begin
  count := 0;
  for i := 0 to 7 do
    if (val and (1 shl i)) <> 0 then Inc(count);
  Result := count mod 2 = 0;
end;
begin
  WriteLn(IsEvenParity($03)); // 00000011 -> 2 ones -> Even
  WriteLn(IsEvenParity($07)); // 00000111 -> 3 ones -> Odd
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
