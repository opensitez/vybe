use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 50: BitSet, BitArray & High-Performance Bitwise Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tbits_class_basic_operations() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var b: TBits;
begin
  b := TBits.Create;
  b.Size := 16;
  b.Bits[5] := True;
  b.Bits[10] := True;
  WriteLn(b.Bits[5]);
  WriteLn(b.Bits[6]);
  WriteLn(b.Bits[10]);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_tbits_openbit_first_clear() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var b: TBits;
begin
  b := TBits.Create;
  b.Size := 8;
  b.Bits[0] := True;
  b.Bits[1] := True;
  WriteLn(b.OpenBit);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_bitmask_set_bit() {
    let out = run_pascal(
        r#"
program Test;
function SetBit(flags, bitPos: Integer): Integer;
begin
  Result := flags or (1 shl bitPos);
end;
begin
  WriteLn(SetBit(0, 3));
end.
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_bitmask_clear_bit() {
    let out = run_pascal(
        r#"
program Test;
function ClearBit(flags, bitPos: Integer): Integer;
begin
  Result := flags and not (1 shl bitPos);
end;
begin
  WriteLn(ClearBit(15, 2));
end.
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_bitmask_toggle_bit() {
    let out = run_pascal(
        r#"
program Test;
function ToggleBit(flags, bitPos: Integer): Integer;
begin
  Result := flags xor (1 shl bitPos);
end;
begin
  WriteLn(ToggleBit(8, 3));
  WriteLn(ToggleBit(0, 3));
end.
"#,
    );
    assert_eq!(out, vec!["0", "8"]);
}

#[test]
fn test_bitmask_test_bit() {
    let out = run_pascal(
        r#"
program Test;
function TestBit(flags, bitPos: Integer): Boolean;
begin
  Result := (flags and (1 shl bitPos)) <> 0;
end;
begin
  WriteLn(TestBit(12, 2));
  WriteLn(TestBit(12, 0));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_popcount_population_count() {
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
  WriteLn(PopCount($FF));
  WriteLn(PopCount($0F0F));
end.
"#,
    );
    assert_eq!(out, vec!["8", "8"]);
}

#[test]
fn test_extract_least_significant_set_bit() {
    let out = run_pascal(
        r#"
program Test;
function ExtractLSB(val: Integer): Integer;
begin
  Result := val and (-val);
end;
begin
  WriteLn(ExtractLSB(12));
end.
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_rotate_left_32bit() {
    let out = run_pascal(
        r#"
program Test;
function RotateLeft(val: DWord; shift: Byte): DWord;
begin
  Result := (val shl shift) or (val shr (32 - shift));
end;
begin
  WriteLn(HexStr(RotateLeft($80000001, 1), 8));
end.
"#,
    );
    assert_eq!(out, vec!["00000003"]);
}

#[test]
fn test_rotate_right_32bit() {
    let out = run_pascal(
        r#"
program Test;
function RotateRight(val: DWord; shift: Byte): DWord;
begin
  Result := (val shr shift) or (val shl (32 - shift));
end;
begin
  WriteLn(HexStr(RotateRight($00000003, 1), 8));
end.
"#,
    );
    assert_eq!(out, vec!["80000001"]);
}

#[test]
fn test_large_set_byte_representation() {
    let out = run_pascal(
        r#"
program Test;
type TByteSet = set of 0..7;
var s: TByteSet;
    b: Byte;
begin
  s := [1, 3, 5];
  Move(s, b, 1);
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_tbits_size_resizing_clears_or_expands() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var b: TBits;
begin
  b := TBits.Create;
  b.Size := 4;
  b.Bits[0] := True; b.Bits[3] := True;
  b.Size := 8;
  WriteLn(b.Bits[0]);
  WriteLn(b.Bits[7]);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tbits_iteration_count_set_bits() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var b: TBits; i, count: Integer;
begin
  b := TBits.Create;
  b.Size := 10;
  b.Bits[1] := True; b.Bits[4] := True; b.Bits[9] := True;
  count := 0;
  for i := 0 to b.Size - 1 do
    if b.Bits[i] then Inc(count);
  WriteLn(count);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_bit_field_struct_simulation() {
    let out = run_pascal(
        r#"
program Test;
type TBitFieldRec = record
  PackedData: Word;
end;

procedure SetFieldA(var r: TBitFieldRec; val: Byte);
begin
  r.PackedData := (r.PackedData and not $000F) or (val and $0F);
end;

function GetFieldA(const r: TBitFieldRec): Byte;
begin
  Result := r.PackedData and $0F;
end;

var r: TBitFieldRec;
begin
  r.PackedData := 0;
  SetFieldA(r, 12);
  WriteLn(GetFieldA(r));
end.
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn test_bitwise_nibble_swap() {
    let out = run_pascal(
        r#"
program Test;
function SwapNibbles(b: Byte): Byte;
begin
  Result := (b shl 4) or (b shr 4);
end;
begin
  WriteLn(HexStr(SwapNibbles($AB), 2));
end.
"#,
    );
    assert_eq!(out, vec!["BA"]);
}

#[test]
fn test_bitwise_int64_masking() {
    let out = run_pascal(
        r#"
program Test;
var val: Int64;
begin
  val := $123456789ABCDEF0;
  WriteLn(HexStr(val and $FFFF, 4));
end.
"#,
    );
    assert_eq!(out, vec!["DEF0"]);
}

#[test]
fn test_is_power_of_two_check() {
    let out = run_pascal(
        r#"
program Test;
function IsPowerOfTwo(n: Integer): Boolean;
begin
  Result := (n > 0) and ((n and (n - 1)) = 0);
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
fn test_tbits_clear_all() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var b: TBits; i: Integer; allClear: Boolean;
begin
  b := TBits.Create;
  b.Size := 5;
  b.Bits[0] := True; b.Bits[2] := True;
  b.Size := 0; b.Size := 5;
  allClear := True;
  for i := 0 to 4 do if b.Bits[i] then allClear := False;
  WriteLn(allClear);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_bitwise_sign_extension_check() {
    let out = run_pascal(
        r#"
program Test;
var sb: ShortInt;
    i: Integer;
begin
  sb := -1;
  i := sb;
  WriteLn(i = -1);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_bit_array_record_wrapper() {
    let out = run_pascal(
        r#"
program Test;
type TBitArray64 = record
  Words: array[0..1] of DWord;
end;

procedure SetBit64(var ba: TBitArray64; bitIdx: Integer);
begin
  ba.Words[bitIdx div 32] := ba.Words[bitIdx div 32] or (1 shl (bitIdx mod 32));
end;

function TestBit64(const ba: TBitArray64; bitIdx: Integer): Boolean;
begin
  Result := (ba.Words[bitIdx div 32] and (1 shl (bitIdx mod 32))) <> 0;
end;

var ba: TBitArray64;
begin
  FillChar(ba, SizeOf(ba), 0);
  SetBit64(ba, 45);
  WriteLn(TestBit64(ba, 45));
  WriteLn(TestBit64(ba, 44));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
