use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 83: Operator Overloading for Custom Record Types
// ═══════════════════════════════════════════════════════════

#[test]
fn test_operator_add_subtract_record() {
    let out = run_pascal(r#"
program Test;
type TPoint = record
  X, Y: Integer;
  class operator Add(a, b: TPoint): TPoint;
  class operator Subtract(a, b: TPoint): TPoint;
end;
class operator TPoint.Add(a, b: TPoint): TPoint;
begin
  Result.X := a.X + b.X; Result.Y := a.Y + b.Y;
end;
class operator TPoint.Subtract(a, b: TPoint): TPoint;
begin
  Result.X := a.X - b.X; Result.Y := a.Y - b.Y;
end;

var p1, p2, resAdd, resSub: TPoint;
begin
  p1.X := 10; p1.Y := 20;
  p2.X := 3;  p2.Y := 5;
  resAdd := p1 + p2;
  resSub := p1 - p2;
  WriteLn(resAdd.X.ToString + ',' + resAdd.Y.ToString);
  WriteLn(resSub.X.ToString + ',' + resSub.Y.ToString);
end.
"#);
    assert_eq!(out, vec!["13,25", "7,15"]);
}

#[test]
fn test_operator_multiply_scalar() {
    let out = run_pascal(r#"
program Test;
type TVec = record
  X, Y: Integer;
  class operator Multiply(v: TVec; scalar: Integer): TVec;
end;
class operator TVec.Multiply(v: TVec; scalar: Integer): TVec;
begin
  Result.X := v.X * scalar; Result.Y := v.Y * scalar;
end;
var v, res: TVec;
begin
  v.X := 4; v.Y := 5;
  res := v * 3;
  WriteLn(res.X.ToString + ',' + res.Y.ToString);
end.
"#);
    assert_eq!(out, vec!["12,15"]);
}

#[test]
fn test_operator_equal_notequal() {
    let out = run_pascal(r#"
program Test;
type TSize = record
  W, H: Integer;
  class operator Equal(a, b: TSize): Boolean;
  class operator NotEqual(a, b: TSize): Boolean;
end;
class operator TSize.Equal(a, b: TSize): Boolean;
begin
  Result := (a.W = b.W) and (a.H = b.H);
end;
class operator TSize.NotEqual(a, b: TSize): Boolean;
begin
  Result := not (a = b);
end;

var s1, s2, s3: TSize;
begin
  s1.W := 100; s1.H := 200;
  s2.W := 100; s2.H := 200;
  s3.W := 50;  s3.H := 200;
  WriteLn(s1 = s2);
  WriteLn(s1 <> s3);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_operator_implicit_conversion() {
    let out = run_pascal(r#"
program Test;
type TMoney = record
  Cents: Integer;
  class operator Implicit(aCents: Integer): TMoney;
end;
class operator TMoney.Implicit(aCents: Integer): TMoney;
begin
  Result.Cents := aCents;
end;
var m: TMoney;
begin
  m := 500; // Implicit assignment from Integer
  WriteLn(m.Cents);
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_operator_explicit_conversion() {
    let out = run_pascal(r#"
program Test;
type TMoney = record
  Cents: Integer;
  class operator Explicit(const m: TMoney): Integer;
end;
class operator TMoney.Explicit(const m: TMoney): Integer;
begin
  Result := m.Cents;
end;
var m: TMoney; val: Integer;
begin
  m.Cents := 750;
  val := Integer(m); // Explicit cast
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["750"]);
}

#[test]
fn test_operator_negative_positive() {
    let out = run_pascal(r#"
program Test;
type TVec = record
  X, Y: Integer;
  class operator Negative(v: TVec): TVec;
  class operator Positive(v: TVec): TVec;
end;
class operator TVec.Negative(v: TVec): TVec;
begin
  Result.X := -v.X; Result.Y := -v.Y;
end;
class operator TVec.Positive(v: TVec): TVec;
begin
  Result := v;
end;

var v, n, p: TVec;
begin
  v.X := 10; v.Y := -20;
  n := -v;
  p := +v;
  WriteLn(n.X.ToString + ',' + n.Y.ToString);
  WriteLn(p.X.ToString + ',' + p.Y.ToString);
end.
"#);
    assert_eq!(out, vec!["-10,20", "10,-20"]);
}

#[test]
fn test_operator_inc_dec() {
    let out = run_pascal(r#"
program Test;
type TCounter = record
  Value: Integer;
  class operator Inc(c: TCounter): TCounter;
  class operator Dec(c: TCounter): TCounter;
end;
class operator TCounter.Inc(c: TCounter): TCounter;
begin
  Result.Value := c.Value + 1;
end;
class operator TCounter.Dec(c: TCounter): TCounter;
begin
  Result.Value := c.Value - 1;
end;

var cnt: TCounter;
begin
  cnt.Value := 10;
  Inc(cnt);
  WriteLn(cnt.Value);
  Dec(cnt);
  WriteLn(cnt.Value);
end.
"#);
    assert_eq!(out, vec!["11", "10"]);
}

#[test]
fn test_operator_greaterthan_lessthan() {
    let out = run_pascal(r#"
program Test;
type TScore = record
  Points: Integer;
  class operator GreaterThan(a, b: TScore): Boolean;
  class operator LessThan(a, b: TScore): Boolean;
end;
class operator TScore.GreaterThan(a, b: TScore): Boolean;
begin
  Result := a.Points > b.Points;
end;
class operator TScore.LessThan(a, b: TScore): Boolean;
begin
  Result := a.Points < b.Points;
end;

var s1, s2: TScore;
begin
  s1.Points := 100; s2.Points := 50;
  WriteLn(s1 > s2);
  WriteLn(s2 < s1);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_operator_divide_scalar() {
    let out = run_pascal(r#"
program Test;
type TVec = record
  X, Y: Integer;
  class operator Divide(v: TVec; d: Integer): TVec;
end;
class operator TVec.Divide(v: TVec; d: Integer): TVec;
begin
  Result.X := v.X div d; Result.Y := v.Y div d;
end;
var v, res: TVec;
begin
  v.X := 100; v.Y := 50;
  res := v / 2;
  WriteLn(res.X.ToString + ',' + res.Y.ToString);
end.
"#);
    assert_eq!(out, vec!["50,25"]);
}

#[test]
fn test_operator_chaining_expression() {
    let out = run_pascal(r#"
program Test;
type TPoint = record
  X: Integer;
  class operator Add(a, b: TPoint): TPoint;
  class operator Subtract(a, b: TPoint): TPoint;
end;
class operator TPoint.Add(a, b: TPoint): TPoint; begin Result.X := a.X + b.X; end;
class operator TPoint.Subtract(a, b: TPoint): TPoint; begin Result.X := a.X - b.X; end;

var p1, p2, p3, res: TPoint;
begin
  p1.X := 10; p2.X := 20; p3.X := 5;
  res := p1 + p2 - p3; // (10 + 20) - 5 = 25
  WriteLn(res.X);
end.
"#);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_operator_complex_number_multiplication() {
    let out = run_pascal(r#"
program Test;
type TComplex = record
  R, I: Integer;
  class operator Multiply(a, b: TComplex): TComplex;
end;
class operator TComplex.Multiply(a, b: TComplex): TComplex;
begin
  // (a+bi)(c+di) = (ac - bd) + (ad + bc)i
  Result.R := (a.R * b.R) - (a.I * b.I);
  Result.I := (a.R * b.I) + (a.I * b.R);
end;
var c1, c2, res: TComplex;
begin
  c1.R := 2; c1.I := 3;
  c2.R := 4; c2.I := 5;
  res := c1 * c2; // (8-15) + (10+12)i = -7 + 22i
  WriteLn(res.R.ToString + '+' + res.I.ToString + 'i');
end.
"#);
    assert_eq!(out, vec!["-7+22i"]);
}

#[test]
fn test_operator_greaterthanorequal_lessthanorequal() {
    let out = run_pascal(r#"
program Test;
type TVal = record
  V: Integer;
  class operator GreaterThanOrEqual(a, b: TVal): Boolean;
  class operator LessThanOrEqual(a, b: TVal): Boolean;
end;
class operator TVal.GreaterThanOrEqual(a, b: TVal): Boolean; begin Result := a.V >= b.V; end;
class operator TVal.LessThanOrEqual(a, b: TVal): Boolean; begin Result := a.V <= b.V; end;

var v1, v2: TVal;
begin
  v1.V := 10; v2.V := 10;
  WriteLn(v1 >= v2);
  WriteLn(v1 <= v2);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_operator_in_custom_record_set() {
    let out = run_pascal(r#"
program Test;
type TRangeRec = record
  MinVal, MaxVal: Integer;
  class operator In(elem: Integer; const r: TRangeRec): Boolean;
end;
class operator TRangeRec.In(elem: Integer; const r: TRangeRec): Boolean;
begin
  Result := (elem >= r.MinVal) and (elem <= r.MaxVal);
end;
var r: TRangeRec;
begin
  r.MinVal := 10; r.MaxVal := 20;
  WriteLn(15 in r);
  WriteLn(25 in r);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_operator_implicit_string_conversion() {
    let out = run_pascal(r#"
program Test;
type TMyStr = record
  Text: String;
  class operator Implicit(const s: String): TMyStr;
  class operator Implicit(const m: TMyStr): String;
end;
class operator TMyStr.Implicit(const s: String): TMyStr; begin Result.Text := s; end;
class operator TMyStr.Implicit(const m: TMyStr): String; begin Result := m.Text; end;

var m: TMyStr; strVal: String;
begin
  m := 'HelloWrapper'; // Implicit string -> TMyStr
  strVal := m;          // Implicit TMyStr -> string
  WriteLn(strVal);
end.
"#);
    assert_eq!(out, vec!["HelloWrapper"]);
}

#[test]
fn test_operator_logical_and_or_not() {
    let out = run_pascal(r#"
program Test;
type TBitFlags = record
  Bits: Byte;
  class operator LogicalAnd(a, b: TBitFlags): TBitFlags;
  class operator LogicalOr(a, b: TBitFlags): TBitFlags;
end;
class operator TBitFlags.LogicalAnd(a, b: TBitFlags): TBitFlags; begin Result.Bits := a.Bits and b.Bits; end;
class operator TBitFlags.LogicalOr(a, b: TBitFlags): TBitFlags; begin Result.Bits := a.Bits or b.Bits; end;

var f1, f2, rAnd, rOr: TBitFlags;
begin
  f1.Bits := $0F; f2.Bits := $33;
  rAnd := f1 and f2; // $0F and $33 = $03
  rOr  := f1 or f2;  // $0F or $33 = $3F
  WriteLn(HexStr(rAnd.Bits, 2));
  WriteLn(HexStr(rOr.Bits, 2));
end.
"#);
    assert_eq!(out, vec!["03", "3F"]);
}

#[test]
fn test_operator_modulus() {
    let out = run_pascal(r#"
program Test;
type TNum = record
  V: Integer;
  class operator Modulus(a, b: TNum): TNum;
end;
class operator TNum.Modulus(a, b: TNum): TNum;
begin
  Result.V := a.V mod b.V;
end;
var n1, n2, res: TNum;
begin
  n1.V := 17; n2.V := 5;
  res := n1 mod n2;
  WriteLn(res.V);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_operator_leftshift_rightshift() {
    let out = run_pascal(r#"
program Test;
type TWordRec = record
  W: Word;
  class operator LeftShift(a: TWordRec; shift: Integer): TWordRec;
  class operator RightShift(a: TWordRec; shift: Integer): TWordRec;
end;
class operator TWordRec.LeftShift(a: TWordRec; shift: Integer): TWordRec; begin Result.W := a.W shl shift; end;
class operator TWordRec.RightShift(a: TWordRec; shift: Integer): TWordRec; begin Result.W := a.W shr shift; end;

var w, rLeft, rRight: TWordRec;
begin
  w.W := 8;
  rLeft := w shl 2;  // 8 << 2 = 32
  rRight := w shr 1; // 8 >> 1 = 4
  WriteLn(rLeft.W);
  WriteLn(rRight.W);
end.
"#);
    assert_eq!(out, vec!["32", "4"]);
}

#[test]
fn test_operator_commutative_overload() {
    let out = run_pascal(r#"
program Test;
type TVec = record
  X: Integer;
  class operator Add(v: TVec; scalar: Integer): TVec;
  class operator Add(scalar: Integer; v: TVec): TVec;
end;
class operator TVec.Add(v: TVec; scalar: Integer): TVec; begin Result.X := v.X + scalar; end;
class operator TVec.Add(scalar: Integer; v: TVec): TVec; begin Result.X := v.X + scalar; end;

var v, r1, r2: TVec;
begin
  v.X := 10;
  r1 := v + 5;
  r2 := 5 + v;
  WriteLn(r1.X = r2.X);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_operator_record_in_array_sum() {
    let out = run_pascal(r#"
program Test;
type TItem = record
  Val: Integer;
  class operator Add(a, b: TItem): TItem;
end;
class operator TItem.Add(a, b: TItem): TItem; begin Result.Val := a.Val + b.Val; end;

var items: array[0..2] of TItem; sum: TItem; i: Integer;
begin
  items[0].Val := 10; items[1].Val := 20; items[2].Val := 30;
  sum.Val := 0;
  for i := 0 to 2 do
    sum := sum + items[i];
  WriteLn(sum.Val);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_operator_intdiv() {
    let out = run_pascal(r#"
program Test;
type TRec = record
  V: Integer;
  class operator IntDivide(a, b: TRec): TRec;
end;
class operator TRec.IntDivide(a, b: TRec): TRec; begin Result.V := a.V div b.V; end;

var r1, r2, res: TRec;
begin
  r1.V := 25; r2.V := 4;
  res := r1 div r2;
  WriteLn(res.V);
end.
"#);
    assert_eq!(out, vec!["6"]);
}
