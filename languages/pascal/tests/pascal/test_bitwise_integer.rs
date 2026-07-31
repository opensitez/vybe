/// Tests for integer bitwise operations in Pascal/Delphi:
/// Bitwise AND, OR, XOR, NOT on integer operands,
/// bit masks, flag patterns, bit shifting combinations.
use super::helpers::run_pascal;

// ===================================================================
// BITWISE AND ON INTEGERS
// ===================================================================

#[test]
fn bitwise_and_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Integer;
begin
  a := 5;
  b := 3;
  WriteLn(a and b);
end."#
        ),
        &["1"]
    );
}

#[test]
fn bitwise_and_mask() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 255;
  WriteLn(n and $0F);
end."#
        ),
        &["15"]
    );
}

#[test]
fn bitwise_and_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(7 and 0);
end."#
        ),
        &["0"]
    );
}

#[test]
fn bitwise_and_self() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(42 and 42);
end."#
        ),
        &["42"]
    );
}

// ===================================================================
// BITWISE OR ON INTEGERS
// ===================================================================

#[test]
fn bitwise_or_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Integer;
begin
  a := 5;
  b := 3;
  WriteLn(a or b);
end."#
        ),
        &["7"]
    );
}

#[test]
fn bitwise_or_with_zero() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(42 or 0);
end."#
        ),
        &["42"]
    );
}

#[test]
fn bitwise_or_flags() {
    assert_eq!(
        run_pascal(
            r#"program T;
const
  FLAG_READ  = 1;
  FLAG_WRITE = 2;
  FLAG_EXEC  = 4;
var perms: Integer;
begin
  perms := FLAG_READ or FLAG_WRITE;
  WriteLn(perms);
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// BITWISE XOR ON INTEGERS
// ===================================================================

#[test]
fn bitwise_xor_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(5 xor 3);
end."#
        ),
        &["6"]
    );
}

#[test]
fn bitwise_xor_same() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn(42 xor 42);
end."#
        ),
        &["0"]
    );
}

#[test]
fn bitwise_xor_swap() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b: Integer;
begin
  a := 10;
  b := 20;
  a := a xor b;
  b := a xor b;
  a := a xor b;
  WriteLn(a);
  WriteLn(b);
end."#
        ),
        &["20", "10"]
    );
}

// ===================================================================
// BIT SHIFTING COMBINATIONS
// ===================================================================

#[test]
fn shl_then_shr() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 1;
  n := n shl 8;
  n := n shr 4;
  WriteLn(n);
end."#
        ),
        &["16"]
    );
}

#[test]
fn shl_power_of_two() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, result: Integer;
begin
  result := 1;
  for i := 1 to 4 do
    result := result shl 1;
  WriteLn(result);
end."#
        ),
        &["16"]
    );
}

#[test]
fn shr_halve() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 64;
  WriteLn(n shr 1);
  WriteLn(n shr 2);
  WriteLn(n shr 3);
end."#
        ),
        &["32", "16", "8"]
    );
}

// ===================================================================
// BIT FLAG PATTERNS
// ===================================================================

#[test]
fn bit_flag_test() {
    assert_eq!(
        run_pascal(
            r#"program T;
const FLAG = 4;
var flags: Integer;
begin
  flags := 7;
  if (flags and FLAG) <> 0 then
    WriteLn('flag set')
  else
    WriteLn('flag clear');
end."#
        ),
        &["flag set"]
    );
}

#[test]
fn bit_flag_set_and_test() {
    assert_eq!(
        run_pascal(
            r#"program T;
var flags: Integer;
begin
  flags := 0;
  flags := flags or 2;
  flags := flags or 8;
  WriteLn(flags);
end."#
        ),
        &["10"]
    );
}

#[test]
fn bit_flag_clear() {
    assert_eq!(
        run_pascal(
            r#"program T;
var flags: Integer;
begin
  flags := 15;
  flags := flags and (not 4);
  WriteLn(flags);
end."#
        ),
        &["11"]
    );
}

// ===================================================================
// COMBINED BITWISE EXPRESSIONS
// ===================================================================

#[test]
fn bitwise_chain() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: Integer;
begin
  a := ($FF and $3F) or $40;
  WriteLn(a);
end."#
        ),
        &["127"]
    );
}

#[test]
fn even_odd_via_and() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 6;
  if (n and 1) = 0 then WriteLn('even') else WriteLn('odd');
  n := 7;
  if (n and 1) = 0 then WriteLn('even') else WriteLn('odd');
end."#
        ),
        &["even", "odd"]
    );
}

#[test]
fn nibble_extraction() {
    assert_eq!(
        run_pascal(
            r#"program T;
var b: Integer;
    hi, lo: Integer;
begin
  b := $AB;
  lo := b and $0F;
  hi := (b shr 4) and $0F;
  WriteLn(lo);
  WriteLn(hi);
end."#
        ),
        &["11", "10"]
    );
}

#[test]
fn bitwise_not_on_byte() {
    assert_eq!(
        run_pascal(
            r#"program T;
var b: Integer;
begin
  b := $FF;
  WriteLn((not b) and $FF);
end."#
        ),
        &["0"]
    );
}

// ── `xor` resolves by OPERAND TYPE, not by spelling ────────────────────────
//
// Pascal has one `xor` token with two meanings: bitwise when both operands are
// integers, logical otherwise. The compiler emits the same bit-xor opcode for
// both and materializes a Boolean only in the second case, which is why this is
// a profile property (`xor_is_logical_for_non_integers`) rather than a choice
// of emit target — see builtinslotplan.md §3i.
//
// The bitwise cases above all use integer LITERALS. These pin the other half:
// the same operator on Boolean operands, and on integer VARIABLES, where the
// decision is made from a type hint rather than the literal's shape.

#[test]
fn xor_on_boolean_literals_is_logical() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true xor false); end."#),
        &["true"]
    );
}

#[test]
fn xor_of_equal_booleans_is_false() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn(true xor true); end."#),
        &["false"]
    );
}

#[test]
fn xor_on_boolean_variables_is_logical() {
    assert_eq!(
        run_pascal(
            r#"program T; var p, q: Boolean; begin p := True; q := False; WriteLn(p xor q); end."#
        ),
        &["true"]
    );
}

#[test]
fn xor_on_integer_variables_is_bitwise() {
    assert_eq!(
        run_pascal(r#"program T; var a, b: Integer; begin a := 12; b := 10; WriteLn(a xor b); end."#),
        &["6"]
    );
}

#[test]
fn xor_on_comparison_results_is_logical() {
    assert_eq!(
        run_pascal(r#"program T; begin WriteLn((5 > 3) xor (3 > 5)); end."#),
        &["true"]
    );
}
