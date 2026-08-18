use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 75: Random Number Generation & Seed Control
// ═══════════════════════════════════════════════════════════

#[test]
fn test_random_float_range() {
    let out = run_pascal(
        r#"
program Test;
var r: Extended;
begin
  RandSeed := 42;
  r := Random;
  WriteLn((r >= 0.0) and (r < 1.0));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_integer_bound() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer;
begin
  RandSeed := 100;
  val := Random(10);
  WriteLn((val >= 0) and (val < 10));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_deterministic_seed_reproducibility() {
    let out = run_pascal(
        r#"
program Test;
var v1, v2: Integer;
begin
  RandSeed := 12345;
  v1 := Random(1000);

  RandSeed := 12345;
  v2 := Random(1000);

  WriteLn(v1 = v2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_math_randomrange() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var val: Integer;
begin
  RandSeed := 999;
  val := RandomRange(50, 100);
  WriteLn((val >= 50) and (val < 100));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_math_randomfrom_array() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var arr: array[0..2] of Integer; val: Integer;
begin
  RandSeed := 555;
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  val := RandomFrom(arr);
  WriteLn((val = 10) or (val = 20) or (val = 30));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_randomize_procedure() {
    let out = run_pascal(
        r#"
program Test;
begin
  Randomize;
  WriteLn(RandSeed <> 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_boolean() {
    let out = run_pascal(
        r#"
program Test;
function RandomBool: Boolean;
begin
  Result := Random(2) = 1;
end;
begin
  RandSeed := 777;
  WriteLn((RandomBool = True) or (RandomBool = False));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_shuffle_array() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[0..3] of Integer; i, j, tmp: Integer;
begin
  RandSeed := 123;
  arr[0] := 1; arr[1] := 2; arr[2] := 3; arr[3] := 4;
  for i := 3 downto 1 do
  begin
    j := Random(i + 1);
    tmp := arr[i]; arr[i] := arr[j]; arr[j] := tmp;
  end;
  WriteLn(arr[0] + arr[1] + arr[2] + arr[3] = 10);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_sequence_distribution() {
    let out = run_pascal(
        r#"
program Test;
var i, val: Integer; inRange: Boolean;
begin
  RandSeed := 888;
  inRange := True;
  for i := 1 to 20 do
  begin
    val := Random(5);
    if (val < 0) or (val >= 5) then inRange := False;
  end;
  WriteLn(inRange);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_negative_randomrange() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var val: Integer;
begin
  RandSeed := 321;
  val := RandomRange(-10, -5);
  WriteLn((val >= -10) and (val < -5));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_range_of_one() {
    let out = run_pascal(
        r#"
program Test;
begin
  RandSeed := 10;
  WriteLn(Random(1) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_randseed_restore() {
    let out = run_pascal(
        r#"
program Test;
var savedSeed: Integer; v1, v2: Integer;
begin
  savedSeed := RandSeed;
  RandSeed := 500;
  v1 := Random(100);
  RandSeed := 500;
  v2 := Random(100);
  RandSeed := savedSeed;
  WriteLn(v1 = v2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_float_scaling() {
    let out = run_pascal(
        r#"
program Test;
var val: Double;
begin
  RandSeed := 456;
  val := Random * 100.0;
  WriteLn((val >= 0.0) and (val < 100.0));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_int64_generation() {
    let out = run_pascal(
        r#"
program Test;
function RandomInt64: Int64;
begin
  Result := Int64(Random(1000000)) * 1000000 + Random(1000000);
end;
begin
  RandSeed := 789;
  WriteLn(RandomInt64 >= 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_char_selection() {
    let out = run_pascal(
        r#"
program Test;
function RandomChar: Char;
begin
  Result := Chr(Ord('A') + Random(26));
end;
begin
  RandSeed := 111;
  WriteLn((RandomChar >= 'A') and (RandomChar <= 'Z'));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_weighted_choice() {
    let out = run_pascal(
        r#"
program Test;
function WeightedChoice(w1, w2: Integer): Integer;
var total, roll: Integer;
begin
  total := w1 + w2;
  roll := Random(total);
  if roll < w1 then Result := 1 else Result := 2;
end;
begin
  RandSeed := 222;
  WriteLn((WeightedChoice(70, 30) = 1) or (WeightedChoice(70, 30) = 2));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_matrix_population() {
    let out = run_pascal(
        r#"
program Test;
var mat: array[0..1, 0..1] of Integer; r, c: Integer; valid: Boolean;
begin
  RandSeed := 333;
  valid := True;
  for r := 0 to 1 do
    for c := 0 to 1 do
    begin
      mat[r, c] := Random(100);
      if (mat[r, c] < 0) or (mat[r, c] >= 100) then valid := False;
    end;
  WriteLn(valid);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_gaussian_approximation() {
    let out = run_pascal(
        r#"
program Test;
function ApproxGaussian: Double;
begin
  Result := (Random + Random + Random) / 3.0;
end;
begin
  RandSeed := 444;
  WriteLn((ApproxGaussian >= 0.0) and (ApproxGaussian <= 1.0));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_enum_selection() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
function RandomColor: TColor;
begin
  Result := TColor(Random(3));
end;
begin
  RandSeed := 555;
  WriteLn(Ord(RandomColor) <= 2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_random_range_same_from_to() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  RandSeed := 666;
  WriteLn(RandomRange(5, 6) = 5);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}
