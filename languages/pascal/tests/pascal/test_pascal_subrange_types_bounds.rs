use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 2: Subrange Types & Bounds Validation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_subrange_integer_assignment() {
    let out = run_pascal(
        r#"
program Test;
type TOneToTen = 1..10;
var x: TOneToTen;
begin
  x := 5;
  WriteLn(x);
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_subrange_character_bounds() {
    let out = run_pascal(
        r#"
program Test;
type TUpperChar = 'A'..'Z';
var c: TUpperChar;
begin
  c := 'M';
  WriteLn(c);
end.
"#,
    );
    assert_eq!(out, vec!["M"]);
}

#[test]
fn test_subrange_as_array_index_type() {
    let out = run_pascal(
        r#"
program Test;
type TIdx = 10..15;
type TArr = array[TIdx] of String;
var a: TArr;
begin
  a[10] := 'First';
  a[15] := 'Last';
  WriteLn(a[10]);
  WriteLn(a[15]);
end.
"#,
    );
    assert_eq!(out, vec!["First", "Last"]);
}

#[test]
fn test_subrange_high_low_functions() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 0..100;
begin
  WriteLn(Low(TScore));
  WriteLn(High(TScore));
end.
"#,
    );
    assert_eq!(out, vec!["0", "100"]);
}

#[test]
fn test_subrange_negative_integers() {
    let out = run_pascal(
        r#"
program Test;
type TRange = -10..10;
var val: TRange;
begin
  val := -5;
  WriteLn(val);
  val := val + 10;
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["-5", "5"]);
}

#[test]
fn test_subrange_enum_subrange() {
    let out = run_pascal(
        r#"
program Test;
type TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
type TWorkDay = Mon..Fri;
var d: TWorkDay;
begin
  d := Wed;
  WriteLn(Ord(d));
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_subrange_for_loop_iteration() {
    let out = run_pascal(
        r#"
program Test;
type TSmall = 1..4;
var i: TSmall;
begin
  for i := Low(TSmall) to High(TSmall) do
    WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn test_subrange_function_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TPercent = 0..100;
function DoublePercent(p: TPercent): Integer;
begin
  Result := p * 2;
end;
begin
  WriteLn(DoublePercent(45));
end.
"#,
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn test_subrange_succ_pred() {
    let out = run_pascal(
        r#"
program Test;
type TRange = 5..15;
var r: TRange;
begin
  r := 10;
  WriteLn(Pred(r));
  WriteLn(Succ(r));
end.
"#,
    );
    assert_eq!(out, vec!["9", "11"]);
}

#[test]
fn test_subrange_inc_dec() {
    let out = run_pascal(
        r#"
program Test;
type TRange = 1..10;
var val: TRange;
begin
  val := 5;
  Inc(val);
  WriteLn(val);
  Dec(val, 2);
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["6", "4"]);
}

#[test]
fn test_subrange_case_statement() {
    let out = run_pascal(
        r#"
program Test;
type TGradeScale = 1..100;
var score: TGradeScale;
begin
  score := 85;
  case score of
    90..100: WriteLn('A');
    80..89: WriteLn('B');
    70..79: WriteLn('C');
  else
    WriteLn('F');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn test_subrange_set_of_subrange() {
    let out = run_pascal(
        r#"
program Test;
type TSmallRange = 1..5;
type TSmallSet = set of TSmallRange;
var s: TSmallSet;
begin
  s := [2, 4];
  WriteLn(2 in s);
  WriteLn(3 in s);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_subrange_record_field() {
    let out = run_pascal(
        r#"
program Test;
type TAge = 0..120;
type TPerson = record
  Name: String;
  Age: TAge;
end;
var p: TPerson;
begin
  p.Name := 'Alice';
  p.Age := 30;
  WriteLn(p.Name);
  WriteLn(p.Age);
end.
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn test_subrange_type_coercion_to_integer() {
    let out = run_pascal(
        r#"
program Test;
type TSub = 1..5;
var s: TSub;
    i: Integer;
begin
  s := 3;
  i := s * 10;
  WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_subrange_byte_subrange() {
    let out = run_pascal(
        r#"
program Test;
type TByteSub = Byte(0)..Byte(100);
var b: TByteSub;
begin
  b := 50;
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_subrange_char_case_conversion() {
    let out = run_pascal(
        r#"
program Test;
type TLowerChar = 'a'..'z';
var lc: TLowerChar;
    uc: Char;
begin
  lc := 'g';
  uc := UpCase(lc);
  WriteLn(uc);
end.
"#,
    );
    assert_eq!(out, vec!["G"]);
}

#[test]
fn test_subrange_odd_even_check() {
    let out = run_pascal(
        r#"
program Test;
type TRange = 1..50;
var num: TRange;
begin
  num := 27;
  WriteLn(Odd(num));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subrange_multi_dimensional_array_bounds() {
    let out = run_pascal(
        r#"
program Test;
type TRow = 1..3;
type TCol = 1..3;
var grid: array[TRow, TCol] of Integer;
begin
  grid[2, 3] := 99;
  WriteLn(grid[2, 3]);
end.
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_subrange_typed_constant() {
    let out = run_pascal(
        r#"
program Test;
type TLevel = 1..5;
const DefaultLevel: TLevel = 3;
begin
  WriteLn(DefaultLevel);
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_subrange_max_int_subrange() {
    let out = run_pascal(
        r#"
program Test;
type TBigSub = 100000..200000;
var val: TBigSub;
begin
  val := 150000;
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["150000"]);
}
