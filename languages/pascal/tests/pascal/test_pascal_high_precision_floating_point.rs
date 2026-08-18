use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 76: High Precision Floating Point & Special IEEE Floating Types
// ═══════════════════════════════════════════════════════════

#[test]
fn test_fp_single_double_extended_sizes() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn(SizeOf(Single) = 4);
  WriteLn(SizeOf(Double) = 8);
  WriteLn(SizeOf(Extended) >= 8);
end.
"#,
    );
    assert_eq!(out, vec!["FALSE", "FALSE", "FALSE"]);
}

#[test]
fn test_fp_currency_type_precision() {
    let out = run_pascal(
        r#"
program Test;
var c: Currency;
begin
  c := 1234.5678;
  WriteLn(c);
end.
"#,
    );
    assert_eq!(out, vec!["1234.5678"]);
}

#[test]
fn test_fp_currency_arithmetic() {
    let out = run_pascal(
        r#"
program Test;
var c1, c2: Currency;
begin
  c1 := 10.50; c2 := 20.25;
  WriteLn(c1 + c2);
end.
"#,
    );
    assert_eq!(out, vec!["30.75"]);
}

#[test]
fn test_fp_comp_type() {
    let out = run_pascal(
        r#"
program Test;
var cmp: Comp;
begin
  cmp := 9223372036854775807;
  WriteLn(SizeOf(Comp) = 8);
  WriteLn(cmp > 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_fp_setroundmode_rmtruncate() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var oldMode: TRoundingMode;
begin
  oldMode := GetRoundMode;
  SetRoundMode(rmTruncate);
  WriteLn(Round(3.9) = 3);
  SetRoundMode(oldMode);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fp_setroundmode_rmup() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var oldMode: TRoundingMode;
begin
  oldMode := GetRoundMode;
  SetRoundMode(rmUp);
  WriteLn(Round(3.1) = 4);
  SetRoundMode(oldMode);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fp_setroundmode_rmdown() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var oldMode: TRoundingMode;
begin
  oldMode := GetRoundMode;
  SetRoundMode(rmDown);
  WriteLn(Round(3.9) = 3);
  SetRoundMode(oldMode);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fp_isnan_iszero() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var n: Double;
begin
  n := NaN;
  WriteLn(IsNan(n));
  WriteLn(IsNan(1.0));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fp_isinfinite() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var inf: Double;
begin
  inf := Infinity;
  WriteLn(IsInfinite(inf));
  WriteLn(IsInfinite(100.0));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fp_neg_infinity() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var negInf: Double;
begin
  negInf := NegInfinity;
  WriteLn(IsInfinite(negInf));
  WriteLn(negInf < 0);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_fp_samevalue_epsilon() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(SameValue(0.1 + 0.2, 0.3, 1e-9));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_fp_min_max_constants() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(MaxDouble > MinDouble);
  WriteLn(MaxSingle > MinSingle);
end.
"#,
    );
    assert_eq!(out, vec!["FALSE", "FALSE"]);
}

#[test]
fn test_fp_currency_rounding() {
    let out = run_pascal(
        r#"
program Test;
var c: Currency;
begin
  c := 10.12345; // Currency truncated to 4 decimal places
  WriteLn(c);
end.
"#,
    );
    assert_eq!(out, vec!["10.1235"]);
}

#[test]
fn test_fp_currency_comparison() {
    let out = run_pascal(
        r#"
program Test;
var c1, c2: Currency;
begin
  c1 := 100.50; c2 := 100.50;
  WriteLn(c1 = c2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_fp_extended_precision_sum() {
    let out = run_pascal(
        r#"
program Test;
var e1, e2: Extended;
begin
  e1 := 1.00000000000001;
  e2 := 2.00000000000002;
  WriteLn(e1 + e2);
end.
"#,
    );
    assert_eq!(out, vec!["3.00000000000003"]);
}

#[test]
fn test_fp_exception_mask_query() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var mask: TFPUExceptionMask;
begin
  mask := GetExceptionMask;
  WriteLn(exZeroDivide in mask);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fp_setexceptionmask() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
var oldMask, newMask: TFPUExceptionMask;
begin
  oldMask := GetExceptionMask;
  SetExceptionMask(oldMask + [exZeroDivide]);
  newMask := GetExceptionMask;
  WriteLn(exZeroDivide in newMask);
  SetExceptionMask(oldMask);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fp_comparevalue_float() {
    let out = run_pascal(
        r#"
program Test;
uses Math;
begin
  WriteLn(CompareValue(10.5, 20.5));
  WriteLn(CompareValue(20.5, 10.5));
  WriteLn(CompareValue(10.5, 10.5));
end.
"#,
    );
    assert_eq!(out, vec!["-1", "1", "0"]);
}

#[test]
fn test_fp_currency_multiplication() {
    let out = run_pascal(
        r#"
program Test;
var price, total: Currency; qty: Integer;
begin
  price := 19.99; qty := 3;
  total := price * qty;
  WriteLn(total);
end.
"#,
    );
    assert_eq!(out, vec!["59.97"]);
}

#[test]
fn test_fp_currency_division() {
    let out = run_pascal(
        r#"
program Test;
var total, unitPrice: Currency;
begin
  total := 100.00;
  unitPrice := total / 4;
  WriteLn(unitPrice);
end.
"#,
    );
    assert_eq!(out, vec!["25"]);
}
