/// Tests for Pascal/Delphi numeric type declarations: Byte, Word, SmallInt,
/// LongInt, Int64, Cardinal, ShortInt, Single, Double — type names as used
/// in real Delphi code, distinct from the basic Integer/Real already tested.
use super::helpers::run_pascal;

// ===================================================================
// BYTE TYPE (0..255)
// ===================================================================

#[test]
fn byte_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var b: Byte;
begin
  b := 200;
  WriteLn(b);
end."#
        ),
        &["200"]
    );
}

#[test]
fn byte_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var b: Byte;
begin
  b := 10;
  WriteLn(b * 2 + 5);
end."#
        ),
        &["25"]
    );
}

#[test]
fn byte_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Double(b: Byte): Integer;
begin
  Result := b * 2;
end;
begin
  WriteLn(Double(100));
end."#
        ),
        &["200"]
    );
}

// ===================================================================
// WORD TYPE (0..65535)
// ===================================================================

#[test]
fn word_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var w: Word;
begin
  w := 1000;
  WriteLn(w);
end."#
        ),
        &["1000"]
    );
}

#[test]
fn word_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var w: Word;
begin
  w := 300;
  WriteLn(w + 200);
end."#
        ),
        &["500"]
    );
}

// ===================================================================
// SMALLINT TYPE (-32768..32767)
// ===================================================================

#[test]
fn smallint_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: SmallInt;
begin
  s := -500;
  WriteLn(s);
end."#
        ),
        &["-500"]
    );
}

#[test]
fn smallint_positive() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: SmallInt;
begin
  s := 1000;
  WriteLn(s * 3);
end."#
        ),
        &["3000"]
    );
}

// ===================================================================
// LONGINT TYPE
// ===================================================================

#[test]
fn longint_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var l: LongInt;
begin
  l := 100000;
  WriteLn(l);
end."#
        ),
        &["100000"]
    );
}

#[test]
fn longint_negative() {
    assert_eq!(
        run_pascal(
            r#"program T;
var l: LongInt;
begin
  l := -100000;
  WriteLn(l);
end."#
        ),
        &["-100000"]
    );
}

// ===================================================================
// INT64 TYPE
// ===================================================================

#[test]
fn int64_large_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Int64;
begin
  n := 1000000000;
  WriteLn(n);
end."#
        ),
        &["1000000000"]
    );
}

#[test]
fn int64_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Int64;
begin
  a := 500000;
  b := 300000;
  WriteLn(a + b);
end."#
        ),
        &["800000"]
    );
}

#[test]
fn int64_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Factorial(n: Int64): Int64;
var i: Int64;
begin
  Result := 1;
  for i := 2 to n do
    Result := Result * i;
end;
begin
  WriteLn(Factorial(10));
end."#
        ),
        &["3628800"]
    );
}

// ===================================================================
// CARDINAL (unsigned 32-bit)
// ===================================================================

#[test]
fn cardinal_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Cardinal;
begin
  c := 42;
  WriteLn(c);
end."#
        ),
        &["42"]
    );
}

// ===================================================================
// SHORTINT TYPE (-128..127)
// ===================================================================

#[test]
fn shortint_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: ShortInt;
begin
  s := -50;
  WriteLn(s);
end."#
        ),
        &["-50"]
    );
}

// ===================================================================
// SINGLE TYPE (32-bit float)
// ===================================================================

#[test]
fn single_float() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: Single;
begin
  s := 1.5;
  WriteLn(s);
end."#
        ),
        &["1.5"]
    );
}

#[test]
fn single_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: Single;
begin
  s := 2.5;
  WriteLn(s * 2.0);
end."#
        ),
        &["5"]
    );
}

// ===================================================================
// DOUBLE TYPE (64-bit float)
// ===================================================================

#[test]
fn double_float() {
    assert_eq!(
        run_pascal(
            r#"program T;
var d: Double;
begin
  d := 3.14;
  WriteLn(d);
end."#
        ),
        &["3.14"]
    );
}

#[test]
fn double_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CircleArea(r: Double): Double;
begin
  Result := 3.14159 * r * r;
end;
var area: Double;
begin
  area := CircleArea(2.0);
  WriteLn(area > 12.0);
end."#
        ),
        &["TRUE"]
    );
}

// ===================================================================
// NUMERIC TYPES IN RECORDS
// ===================================================================

#[test]
fn record_with_numeric_types() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMetrics = record
    Count: Cardinal;
    Total: Int64;
    Average: Double;
  end;
var m: TMetrics;
begin
  m.Count := 3;
  m.Total := 300;
  m.Average := 100.0;
  WriteLn(m.Count);
  WriteLn(m.Total);
  WriteLn(m.Average);
end."#
        ),
        &["3", "300", "100"]
    );
}

// ===================================================================
// NUMERIC TYPE COMPARISONS
// ===================================================================

#[test]
fn byte_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
var b: Byte;
begin
  b := 100;
  if b > 50 then WriteLn('big') else WriteLn('small');
end."#
        ),
        &["big"]
    );
}

#[test]
fn int64_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Int64;
begin
  a := 1000000;
  b := 2000000;
  if a < b then WriteLn('less') else WriteLn('not less');
end."#
        ),
        &["less"]
    );
}
