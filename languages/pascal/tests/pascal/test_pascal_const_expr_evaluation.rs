use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 1: Constant Expressions & Static Evaluation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_const_integer_arithmetic_evaluation() {
    let out = run_pascal(
        r#"
program Test;
const A = 10 + 20 * 3;
begin
  WriteLn(A);
end.
"#,
    );
    assert_eq!(out, vec!["70"]);
}

#[test]
fn test_const_floating_point_division() {
    let out = run_pascal(
        r#"
program Test;
const X = 15.0 / 2.0;
begin
  WriteLn(X);
end.
"#,
    );
    assert_eq!(out, vec!["7.5"]);
}

#[test]
fn test_const_string_concatenation() {
    let out = run_pascal(
        r#"
program Test;
const S1 = 'Hello ';
const S2 = 'World!';
const S3 = S1 + S2;
begin
  WriteLn(S3);
end.
"#,
    );
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn test_const_boolean_logical_expressions() {
    let out = run_pascal(
        r#"
program Test;
const B1 = True and False;
const B2 = True or False;
const B3 = not False;
begin
  WriteLn(B1);
  WriteLn(B2);
  WriteLn(B3);
end.
"#,
    );
    assert_eq!(out, vec!["False", "True", "True"]);
}

#[test]
fn test_const_ord_and_chr_functions() {
    let out = run_pascal(
        r#"
program Test;
const C = 'A';
const Code = Ord(C);
const NextC = Chr(Code + 1);
begin
  WriteLn(Code);
  WriteLn(NextC);
end.
"#,
    );
    assert_eq!(out, vec!["65", "B"]);
}

#[test]
fn test_const_bitwise_shifts_and_masking() {
    let out = run_pascal(
        r#"
program Test;
const Mask = (1 shl 4) or (1 shl 1);
const Shifted = Mask shr 1;
begin
  WriteLn(Mask);
  WriteLn(Shifted);
end.
"#,
    );
    assert_eq!(out, vec!["18", "9"]);
}

#[test]
fn test_const_typed_versus_untyped() {
    let out = run_pascal(
        r#"
program Test;
const UntypedVal = 100;
const TypedVal: Integer = 200;
begin
  WriteLn(UntypedVal + TypedVal);
end.
"#,
    );
    assert_eq!(out, vec!["300"]);
}

#[test]
fn test_const_expr_in_array_bounds() {
    let out = run_pascal(
        r#"
program Test;
const MinIndex = 1;
const MaxIndex = 5;
type TArr = array[MinIndex..MaxIndex] of Integer;
var arr: TArr;
begin
  arr[MinIndex] := 10;
  arr[MaxIndex] := 50;
  WriteLn(arr[MinIndex]);
  WriteLn(arr[MaxIndex]);
end.
"#,
    );
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn test_const_nested_references() {
    let out = run_pascal(
        r#"
program Test;
const Base = 10;
const Level1 = Base * 2;
const Level2 = Level1 + 5;
const Level3 = Level2 * Level1;
begin
  WriteLn(Level3);
end.
"#,
    );
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_const_sizeof_evaluation() {
    let out = run_pascal(
        r#"
program Test;
const IntSize = SizeOf(Integer);
begin
  WriteLn(IntSize > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_const_pred_succ_evaluation() {
    let out = run_pascal(
        r#"
program Test;
const StartVal = 10;
const PrevVal = Pred(StartVal);
const NextVal = Succ(StartVal);
begin
  WriteLn(PrevVal);
  WriteLn(NextVal);
end.
"#,
    );
    assert_eq!(out, vec!["9", "11"]);
}

#[test]
fn test_const_hexadecimal_values() {
    let out = run_pascal(
        r#"
program Test;
const HexVal = $FF;
const HexAdded = HexVal + $01;
begin
  WriteLn(HexVal);
  WriteLn(HexAdded);
end.
"#,
    );
    assert_eq!(out, vec!["255", "256"]);
}

#[test]
fn test_const_operator_precedence() {
    let out = run_pascal(
        r#"
program Test;
const V1 = 2 + 3 * 4;
const V2 = (2 + 3) * 4;
begin
  WriteLn(V1);
  WriteLn(V2);
end.
"#,
    );
    assert_eq!(out, vec!["14", "20"]);
}

#[test]
fn test_const_subrange_bound_expression() {
    let out = run_pascal(
        r#"
program Test;
const Offset = 5;
type TSub = Offset..(Offset + 10);
var x: TSub;
begin
  x := 7;
  WriteLn(x);
end.
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_const_negation_unary_minus() {
    let out = run_pascal(
        r#"
program Test;
const PosNum = 42;
const NegNum = -PosNum;
const InvNum = -(-PosNum);
begin
  WriteLn(NegNum);
  WriteLn(InvNum);
end.
"#,
    );
    assert_eq!(out, vec!["-42", "42"]);
}

#[test]
fn test_const_relational_comparisons() {
    let out = run_pascal(
        r#"
program Test;
const IsGreater = 10 > 5;
const IsEqual = 10 = 10;
const IsNotEqual = 5 <> 5;
begin
  WriteLn(IsGreater);
  WriteLn(IsEqual);
  WriteLn(IsNotEqual);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_const_low_high_enum_constants() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (Red, Green, Blue);
const FirstColor = Low(TColor);
const LastColor = High(TColor);
begin
  WriteLn(Ord(FirstColor));
  WriteLn(Ord(LastColor));
end.
"#,
    );
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn test_const_character_literals_and_ordinal() {
    let out = run_pascal(
        r#"
program Test;
const TabChar = #9;
const NewLineChar = #10;
const LetterA = #65;
begin
  WriteLn(LetterA);
end.
"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn test_const_string_length_expression() {
    let out = run_pascal(
        r#"
program Test;
const SampleStr = 'Pascal Programming';
const StrLen = Length(SampleStr);
begin
  WriteLn(StrLen);
end.
"#,
    );
    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_const_modulo_arithmetic() {
    let out = run_pascal(
        r#"
program Test;
const Remainder = 29 mod 7;
begin
  WriteLn(Remainder);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}
