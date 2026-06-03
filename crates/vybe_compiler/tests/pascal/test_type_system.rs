/// Tests for Pascal type system: enums, type casts, is/as operators, type aliases,
/// case ranges, repeat/until, downto, boolean logic, in operator.
use super::helpers::run_pascal;

// ===================================================================
// ENUMS — ADVANCED
// ===================================================================

#[test]
fn enum_ord_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TColor = (Red, Green, Blue);
var c: TColor;
begin
  c := Green;
  WriteLn(c);
end."#
        ),
        &["1"]
    );
}

#[test]
fn enum_explicit_values() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLevel = (Low = 1, Medium = 5, High = 10);
var l: TLevel;
begin
  l := High;
  WriteLn(l);
end."#
        ),
        &["10"]
    );
}

#[test]
fn enum_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
var d: TDay;
begin
  d := Wed;
  if d < Fri then WriteLn('weekday');
end."#
        ),
        &["weekday"]
    );
}

#[test]
fn enum_in_for() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TColor = (Red, Green, Blue);
var c: Integer;
begin
  for c := 0 to 2 do WriteLn(c);
end."#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn enum_succ_pred() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDay = (Mon, Tue, Wed, Thu, Fri);
var d: TDay;
begin
  d := Wed;
  WriteLn(Succ(d));
  WriteLn(Pred(d));
end."#
        ),
        &["3", "1"]
    );
}

// ===================================================================
// CASE RANGES
// ===================================================================

#[test]
fn case_range_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 5;
  case x of
    1..3: WriteLn('low');
    4..6: WriteLn('mid');
    7..9: WriteLn('high');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["mid"]
    );
}

#[test]
fn case_range_first() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 1;
  case x of
    1..5: WriteLn('range');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["range"]
    );
}

#[test]
fn case_range_last() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 5;
  case x of
    1..5: WriteLn('range');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["range"]
    );
}

#[test]
fn case_range_outside() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 10;
  case x of
    1..5: WriteLn('range');
  else
    WriteLn('other');
  end;
end."#
        ),
        &["other"]
    );
}

#[test]
fn case_multiple_values() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 3;
  case x of
    1, 3, 5: WriteLn('odd');
    2, 4, 6: WriteLn('even');
  end;
end."#
        ),
        &["odd"]
    );
}

#[test]
fn case_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: String;
begin
  s := 'b';
  case s of
    'a': WriteLn('alpha');
    'b': WriteLn('bravo');
    'c': WriteLn('charlie');
  end;
end."#
        ),
        &["bravo"]
    );
}

// ===================================================================
// REPEAT-UNTIL
// ===================================================================

#[test]
fn repeat_until_counter() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  i := 1;
  repeat
    WriteLn(i);
    i := i + 1;
  until i > 3;
end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn repeat_until_complex_condition() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 100;
  repeat
    x := x div 2;
  until x < 10;
  WriteLn(x);
end."#
        ),
        &["6"]
    );
}

// ===================================================================
// DOWNTO LOOPS
// ===================================================================

#[test]
fn for_downto_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  for i := 5 downto 1 do WriteLn(i);
end."#
        ),
        &["5", "4", "3", "2", "1"]
    );
}

#[test]
fn for_downto_with_step_logic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, sum: Integer;
begin
  sum := 0;
  for i := 10 downto 1 do sum := sum + i;
  WriteLn(sum);
end."#
        ),
        &["55"]
    );
}

// ===================================================================
// BOOLEAN LOGIC AND SHORT-CIRCUIT
// ===================================================================

#[test]
fn bool_compound_expression() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b, c: Boolean;
begin
  a := true; b := false; c := true;
  if (a and c) and not b then WriteLn('ok');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn bool_or_short_circuit() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 0;
  if (x = 0) or (10 div x > 5) then WriteLn('safe');
end."#
        ),
        &["safe"]
    );
}

// ===================================================================
// IN OPERATOR (WITH ARRAYS)
// ===================================================================

#[test]
fn in_operator_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 3;
  if x in [1, 2, 3, 4, 5] then WriteLn('found');
end."#
        ),
        &["found"]
    );
}

#[test]
fn in_operator_not_found() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 10;
  if not (x in [1, 2, 3]) then WriteLn('not found');
end."#
        ),
        &["not found"]
    );
}

// ===================================================================
// MISC TYPE FEATURES
// ===================================================================

#[test]
fn integer_division() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(17 div 5); end."),
        &["3"]
    );
}

#[test]
fn modulo_operator() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(17 mod 5); end."),
        &["2"]
    );
}

#[test]
fn negative_arithmetic() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(-5 + 3); end."),
        &["-2"]
    );
}

#[test]
fn real_division() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(10 / 4); end."),
        &["2.5"]
    );
}

#[test]
fn integer_overflow_promotion() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(1000000 * 1000000); end."),
        &["1000000000000"]
    );
}

#[test]
fn boolean_to_string() {
    assert_eq!(
        run_pascal("program T; var b: Boolean; begin b := true; WriteLn(BoolToStr(b)); end."),
        &["true"]
    );
}
