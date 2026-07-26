use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 96: Subrange Bounds Checking & Integer Overflow (R+ & Q+)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_subrange_valid_assignment() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
var sub: 1..10;
begin
  sub := 5;
  WriteLn(sub);
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_subrange_upper_bound_overflow_erangeerror() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure TestOverflow;
var sub: 1..5;
begin
  sub := 10;
  WriteLn(sub);
end;
begin
  try
    TestOverflow;
  except
    on E: ERangeError do WriteLn('UpperRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["UpperRangeErrorCaught"]);
}

#[test]
fn test_subrange_lower_bound_underflow_erangeerror() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure TestUnderflow;
var sub: 5..10;
begin
  sub := 2;
  WriteLn(sub);
end;
begin
  try
    TestUnderflow;
  except
    on E: ERangeError do WriteLn('LowerRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["LowerRangeErrorCaught"]);
}

#[test]
fn test_integer_overflow_eintoverflow() {
    let out = run_pascal(
        r#"
program Test;
{$Q+}
uses SysUtils;
procedure TestIntOverflow;
var b: Byte;
begin
  b := 255;
  Inc(b);
  WriteLn(b);
end;
begin
  try
    TestIntOverflow;
  except
    on E: EIntOverflow do WriteLn('IntOverflowCaught');
    on E: ERangeError do WriteLn('IntOverflowCaughtRange');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["IntOverflowCaught"]);
}

#[test]
fn test_subrange_disabled_range_checks() {
    let out = run_pascal(
        r#"
program Test;
{$R-} // Range checks OFF
var sub: 1..5;
begin
  sub := 10; // Allowed without exception when R-
  WriteLn('NoRangeException');
end.
"#,
    );
    assert_eq!(out, vec!["NoRangeException"]);
}

#[test]
fn test_overflow_disabled_overflow_checks() {
    let out = run_pascal(
        r#"
program Test;
{$Q-} // Overflow checks OFF
var b: Byte;
begin
  b := 255;
  b := b + 1; // Wraps around to 0
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_subrange_inc_dec_boundary_check() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
procedure IncBoundary;
var sub: 1..3;
begin
  sub := 3;
  Inc(sub);
end;
begin
  try
    IncBoundary;
  except
    on E: ERangeError do WriteLn('IncRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["IncRangeErrorCaught"]);
}

#[test]
fn test_subrange_enum_bounds_check() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
type TColor = (cRed, cGreen, cBlue);
type TSubColor = cRed..cGreen;
procedure EnumOverflow;
var sc: TSubColor;
begin
  sc := cBlue;
end;
begin
  try
    EnumOverflow;
  except
    on E: ERangeError do WriteLn('EnumRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["EnumRangeErrorCaught"]);
}

#[test]
fn test_subrange_array_indexing_check() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
var arr: array[1..3] of Integer; idx: Integer;
procedure IndexOverflow;
begin
  idx := 5;
  arr[idx] := 100;
end;
begin
  try
    IndexOverflow;
  except
    on E: ERangeError do WriteLn('ArrayIndexErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ArrayIndexErrorCaught"]);
}

#[test]
fn test_subrange_loop_iteration_safety() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
var sub: 1..5; i: Integer;
begin
  for i := 1 to 5 do
  begin
    sub := i;
  end;
  WriteLn(sub);
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_subrange_negative_bounds() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
var sub: -10..-5;
begin
  sub := -7;
  WriteLn(sub);
end.
"#,
    );
    assert_eq!(out, vec!["-7"]);
}

#[test]
fn test_subrange_record_field_check() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
type TRec = record
  Score: 0..100;
end;
procedure SetInvalidScore;
var r: TRec;
begin
  r.Score := 150;
end;
begin
  try
    SetInvalidScore;
  except
    on E: ERangeError do WriteLn('RecordRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RecordRangeErrorCaught"]);
}

#[test]
fn test_overflow_multiplication_check() {
    let out = run_pascal(
        r#"
program Test;
{$Q+}
uses SysUtils;
procedure MulOverflow;
var w: Word;
begin
  w := 1000;
  w := w * 100; // 100,000 > 65,535
end;
begin
  try
    MulOverflow;
  except
    on E: EIntOverflow do WriteLn('MulOverflowCaught');
    on E: ERangeError do WriteLn('MulOverflowCaughtRange');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["MulOverflowCaught"]);
}

#[test]
fn test_subrange_type_conversion_cast() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
type TSmall = 1..10;
var s: TSmall; val: Integer;
begin
  val := 8;
  s := TSmall(val);
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_subrange_property_setter_check() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
type TPercentObj = class
  private FVal: 0..100;
  public property Val: 0..100 read FVal write FVal;
end;
var obj: TPercentObj;
begin
  obj := TPercentObj.Create;
  try
    obj.Val := 200;
  except
    on E: ERangeError do WriteLn('PropertyRangeErrorCaught');
  end;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["PropertyRangeErrorCaught"]);
}

#[test]
fn test_overflow_int64_no_overflow() {
    let out = run_pascal(
        r#"
program Test;
{$Q+}
var a, b: Int64;
begin
  a := 1000000000;
  b := a * 10;
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["10000000000"]);
}

#[test]
fn test_subrange_char_bounds() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
var subChar: 'A'..'Z';
begin
  subChar := 'M';
  WriteLn(subChar);
end.
"#,
    );
    assert_eq!(out, vec!["M"]);
}

#[test]
fn test_subrange_char_overflow() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
procedure CharOverflow;
var subChar: 'A'..'Z';
begin
  subChar := 'a';
end;
begin
  try
    CharOverflow;
  except
    on E: ERangeError do WriteLn('CharRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CharRangeErrorCaught"]);
}

#[test]
fn test_subrange_pred_underflow() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
procedure PredUnderflow;
var sub: 5..10;
begin
  sub := 5;
  sub := Pred(sub);
end;
begin
  try
    PredUnderflow;
  except
    on E: ERangeError do WriteLn('PredRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["PredRangeErrorCaught"]);
}

#[test]
fn test_subrange_succ_overflow() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;
procedure SuccOverflow;
var sub: 1..5;
begin
  sub := 5;
  sub := Succ(sub);
end;
begin
  try
    SuccOverflow;
  except
    on E: ERangeError do WriteLn('SuccRangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["SuccRangeErrorCaught"]);
}
