/// Operator overloads on records: div, mod, inc, dec, and comparison operators.
use super::helpers::run_pascal;

#[test]
fn overload_div_floor_on_pair() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPair = record
    Num, Den: Integer;
    class operator IntDivide(a, b: TPair): TPair;
  end;
class operator TPair.IntDivide(a, b: TPair): TPair;
begin Result.Num := a.Num div b.Den; Result.Den := 1; end;
var p, q, r: TPair;
begin
  p.Num := 17; p.Den := 1;
  q.Num := 1; q.Den := 5;
  r := p div q;
  WriteLn(r.Num);
end."#
        ),
        &["3"]
    );
}

#[test]
fn overload_mod_on_wrapped_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMod = record
    V: Integer;
    class operator Modulus(a: TMod; k: Integer): Integer;
  end;
class operator TMod.Modulus(a: TMod; k: Integer): Integer;
begin Result := a.V mod k; end;
var m: TMod;
begin
  m.V := 23;
  WriteLn(m mod 5);
end."#
        ),
        &["3"]
    );
}

#[test]
fn overload_less_than_on_scores() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TScore = record
    Points: Integer;
    class operator LessThan(a, b: TScore): Boolean;
  end;
class operator TScore.LessThan(a, b: TScore): Boolean;
begin Result := a.Points < b.Points; end;
var a, b: TScore;
begin
  a.Points := 3; b.Points := 9;
  if a < b then WriteLn('lt') else WriteLn('ge');
end."#
        ),
        &["lt"]
    );
}

#[test]
fn overload_less_equal_on_scores() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TScore = record
    Points: Integer;
    class operator LessThanOrEqual(a, b: TScore): Boolean;
  end;
class operator TScore.LessThanOrEqual(a, b: TScore): Boolean;
begin Result := a.Points <= b.Points; end;
var a, b: TScore;
begin
  a.Points := 5; b.Points := 5;
  if a <= b then WriteLn('le') else WriteLn('gt');
end."#
        ),
        &["le"]
    );
}

#[test]
fn overload_equal_on_coordinates() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCoord = record
    X, Y: Integer;
    class operator Equal(a, b: TCoord): Boolean;
  end;
class operator TCoord.Equal(a, b: TCoord): Boolean;
begin Result := (a.X = b.X) and (a.Y = b.Y); end;
var p, q: TCoord;
begin
  p.X := 2; p.Y := 3;
  q.X := 2; q.Y := 3;
  if p = q then WriteLn('eq') else WriteLn('ne');
end."#
        ),
        &["eq"]
    );
}

#[test]
fn overload_not_equal_on_coordinates() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCoord = record
    X: Integer;
    class operator NotEqual(a, b: TCoord): Boolean;
  end;
class operator TCoord.NotEqual(a, b: TCoord): Boolean;
begin Result := a.X <> b.X; end;
var p, q: TCoord;
begin
  p.X := 1; q.X := 2;
  if p <> q then WriteLn('ne') else WriteLn('eq');
end."#
        ),
        &["ne"]
    );
}

#[test]
fn overload_greater_equal_on_magnitude() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMag = record
    V: Integer;
    class operator GreaterThanOrEqual(a, b: TMag): Boolean;
  end;
class operator TMag.GreaterThanOrEqual(a, b: TMag): Boolean;
begin Result := a.V >= b.V; end;
var a, b: TMag;
begin
  a.V := 8; b.V := 3;
  if a >= b then WriteLn('ge') else WriteLn('lt');
end."#
        ),
        &["ge"]
    );
}

#[test]
fn overload_inc_with_step_amount() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStep = record
    N: Integer;
    class operator Inc(var s: TStep; k: Integer);
  end;
class operator TStep.Inc(var s: TStep; k: Integer);
begin s.N := s.N + k; end;
var s: TStep;
begin
  s.N := 10;
  Inc(s, 4);
  WriteLn(s.N);
end."#
        ),
        &["14"]
    );
}

#[test]
fn overload_dec_with_step_amount() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStep = record
    N: Integer;
    class operator Dec(var s: TStep; k: Integer);
  end;
class operator TStep.Dec(var s: TStep; k: Integer);
begin s.N := s.N - k; end;
var s: TStep;
begin
  s.N := 20;
  Dec(s, 6);
  WriteLn(s.N);
end."#
        ),
        &["14"]
    );
}

#[test]
fn overload_negate_on_signed_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TSigned = record
    V: Integer;
    class operator Negative(s: TSigned): TSigned;
  end;
class operator TSigned.Negative(s: TSigned): TSigned;
begin Result.V := -s.V; end;
var s, t: TSigned;
begin
  s.V := 15;
  t := -s;
  WriteLn(t.V);
end."#
        ),
        &["-15"]
    );
}

#[test]
fn overload_implicit_from_integer_wrapper() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TWrap = record
    V: Integer;
    class operator Implicit(n: Integer): TWrap;
  end;
class operator TWrap.Implicit(n: Integer): TWrap;
begin Result.V := n; end;
var w: TWrap;
begin
  w := 42;
  WriteLn(w.V);
end."#
        ),
        &["42"]
    );
}

#[test]
fn overload_div_by_scalar_on_vector() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TVec = record
    X: Integer;
    class operator IntDivide(v: TVec; k: Integer): TVec;
  end;
class operator TVec.IntDivide(v: TVec; k: Integer): TVec;
begin Result.X := v.X div k; end;
var v, w: TVec;
begin
  v.X := 27;
  w := v div 3;
  WriteLn(w.X);
end."#
        ),
        &["9"]
    );
}

#[test]
fn overload_mod_record_pair_remainder() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TDiv = record
  A, B: Integer;
  class operator Modulus(a, b: TDiv): Integer;
  end;
class operator TDiv.Modulus(a, b: TDiv): Integer;
begin Result := a.A mod b.B; end;
var x, y: TDiv;
begin
  x.A := 29; y.B := 7;
  WriteLn(x mod y);
end."#
        ),
        &["1"]
    );
}

#[test]
fn overload_compare_chain_orders_two_items() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TItem = record
    Key: Integer;
    class operator LessThan(a, b: TItem): Boolean;
    class operator GreaterThan(a, b: TItem): Boolean;
  end;
class operator TItem.LessThan(a, b: TItem): Boolean;
begin Result := a.Key < b.Key; end;
class operator TItem.GreaterThan(a, b: TItem): Boolean;
begin Result := a.Key > b.Key; end;
procedure Swap(var a, b: TItem);
var t: TItem;
begin t := a; a := b; b := t; end;
var a, b: TItem;
begin
  a.Key := 9; b.Key := 2;
  if a > b then Swap(a, b);
  WriteLn(a.Key); WriteLn(b.Key);
end."#
        ),
        &["2", "9"]
    );
}

#[test]
fn overload_equal_detects_duplicate_keys() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TKey = record
    Id: Integer;
    class operator Equal(a, b: TKey): Boolean;
  end;
class operator TKey.Equal(a, b: TKey): Boolean;
begin Result := a.Id = b.Id; end;
var a, b: TKey; dup: Boolean;
begin
  a.Id := 7; b.Id := 7;
  dup := a = b;
  if dup then WriteLn('dup') else WriteLn('uniq');
end."#
        ),
        &["dup"]
    );
}

#[test]
fn overload_bitwise_and_on_flag_set() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFlags = record
    Bits: Integer;
    class operator BitwiseAnd(a, b: TFlags): TFlags;
  end;
class operator TFlags.BitwiseAnd(a, b: TFlags): TFlags;
begin Result.Bits := a.Bits and b.Bits; end;
var a, b, c: TFlags;
begin
  a.Bits := 15; b.Bits := 10;
  c := a and b;
  WriteLn(c.Bits);
end."#
        ),
        &["10"]
    );
}

#[test]
fn overload_bitwise_or_on_flag_set() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFlags = record
    Bits: Integer;
    class operator BitwiseOr(a, b: TFlags): TFlags;
  end;
class operator TFlags.BitwiseOr(a, b: TFlags): TFlags;
begin Result.Bits := a.Bits or b.Bits; end;
var a, b, c: TFlags;
begin
  a.Bits := 5; b.Bits := 2;
  c := a or b;
  WriteLn(c.Bits);
end."#
        ),
        &["7"]
    );
}

#[test]
fn overload_bitwise_xor_on_flag_set() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFlags = record
    Bits: Integer;
    class operator BitwiseXor(a, b: TFlags): TFlags;
  end;
class operator TFlags.BitwiseXor(a, b: TFlags): TFlags;
begin Result.Bits := a.Bits xor b.Bits; end;
var a, b, c: TFlags;
begin
  a.Bits := 12; b.Bits := 10;
  c := a xor b;
  WriteLn(c.Bits);
end."#
        ),
        &["6"]
    );
}

#[test]
fn overload_shl_on_shiftable() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TShift = record
    V: Integer;
    class operator LeftShift(s: TShift; n: Integer): TShift;
  end;
class operator TShift.LeftShift(s: TShift; n: Integer): TShift;
begin Result.V := s.V shl n; end;
var s, t: TShift;
begin
  s.V := 3;
  t := s shl 2;
  WriteLn(t.V);
end."#
        ),
        &["12"]
    );
}

#[test]
fn overload_shr_on_shiftable() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TShift = record
    V: Integer;
    class operator RightShift(s: TShift; n: Integer): TShift;
  end;
class operator TShift.RightShift(s: TShift; n: Integer): TShift;
begin Result.V := s.V shr n; end;
var s, t: TShift;
begin
  s.V := 32;
  t := s shr 3;
  WriteLn(t.V);
end."#
        ),
        &["4"]
    );
}

#[test]
fn overload_inc_twice_on_counter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCounter = record
    N: Integer;
    class operator Inc(var c: TCounter);
  end;
class operator TCounter.Inc(var c: TCounter);
begin c.N := c.N + 1; end;
var c: TCounter;
begin
  c.N := 0;
  Inc(c); Inc(c);
  WriteLn(c.N);
end."#
        ),
        &["2"]
    );
}

#[test]
fn overload_dec_twice_on_counter() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCounter = record
    N: Integer;
    class operator Dec(var c: TCounter);
  end;
class operator TCounter.Dec(var c: TCounter);
begin c.N := c.N - 1; end;
var c: TCounter;
begin
  c.N := 5;
  Dec(c); Dec(c);
  WriteLn(c.N);
end."#
        ),
        &["3"]
    );
}

#[test]
fn overload_compare_in_if_else_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TRank = record
    Level: Integer;
    class operator GreaterThan(a, b: TRank): Boolean;
  end;
class operator TRank.GreaterThan(a, b: TRank): Boolean;
begin Result := a.Level > b.Level; end;
var a, b: TRank;
begin
  a.Level := 2; b.Level := 5;
  if a > b then WriteLn('high') else WriteLn('low');
end."#
        ),
        &["low"]
    );
}

#[test]
fn overload_sort_two_records_by_less_than() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TVal = record
    N: Integer;
    class operator LessThan(a, b: TVal): Boolean;
  end;
class operator TVal.LessThan(a, b: TVal): Boolean;
begin Result := a.N < b.N; end;
var a, b: TVal;
begin
  a.N := 30; b.N := 12;
  if b < a then WriteLn(b.N) else WriteLn(a.N);
  if a > b then WriteLn(a.N) else WriteLn(b.N);
end."#
        ),
        &["12", "30"]
    );
}

#[test]
fn overload_div_even_split() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TEven = record
    V: Integer;
    class operator IntDivide(e: TEven; k: Integer): Integer;
  end;
class operator TEven.IntDivide(e: TEven; k: Integer): Integer;
begin Result := e.V div k; end;
var e: TEven;
begin
  e.V := 48;
  WriteLn(e div 6);
end."#
        ),
        &["8"]
    );
}

#[test]
fn overload_mod_odd_remainder() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TOdd = record
    V: Integer;
    class operator Modulus(o: TOdd; k: Integer): Integer;
  end;
class operator TOdd.Modulus(o: TOdd; k: Integer): Integer;
begin Result := o.V mod k; end;
var o: TOdd;
begin
  o.V := 31;
  WriteLn(o mod 4);
end."#
        ),
        &["3"]
    );
}

#[test]
fn overload_negate_then_add_records() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TInt = record
    V: Integer;
    class operator Negative(i: TInt): TInt;
    class operator Add(a, b: TInt): TInt;
  end;
class operator TInt.Negative(i: TInt): TInt;
begin Result.V := -i.V; end;
class operator TInt.Add(a, b: TInt): TInt;
begin Result.V := a.V + b.V; end;
var a, b, c: TInt;
begin
  a.V := 10;
  b := -a;
  c.V := 5;
  WriteLn((b + c).V);
end."#
        ),
        &["-5"]
    );
}

#[test]
fn overload_implicit_then_compare_equal() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBox = record
    V: Integer;
    class operator Implicit(n: Integer): TBox;
    class operator Equal(a, b: TBox): Boolean;
  end;
class operator TBox.Implicit(n: Integer): TBox;
begin Result.V := n; end;
class operator TBox.Equal(a, b: TBox): Boolean;
begin Result := a.V = b.V; end;
var a, b: TBox;
begin
  a := 99;
  b.V := 99;
  if a = b then WriteLn('match') else WriteLn('miss');
end."#
        ),
        &["match"]
    );
}

#[test]
fn overload_greater_than_false_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAge = record
    Years: Integer;
    class operator GreaterThan(a, b: TAge): Boolean;
  end;
class operator TAge.GreaterThan(a, b: TAge): Boolean;
begin Result := a.Years > b.Years; end;
var child, adult: TAge;
begin
  child.Years := 8; adult.Years := 30;
  if child > adult then WriteLn('child') else WriteLn('adult');
end."#
        ),
        &["adult"]
    );
}

#[test]
fn overload_less_equal_on_boundary_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TLimit = record
    Max: Integer;
    class operator LessThanOrEqual(a, b: TLimit): Boolean;
  end;
class operator TLimit.LessThanOrEqual(a, b: TLimit): Boolean;
begin Result := a.Max <= b.Max; end;
var a, b: TLimit;
begin
  a.Max := 100; b.Max := 100;
  if a <= b then WriteLn('ok') else WriteLn('over');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn overload_greater_equal_on_boundary_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TLimit = record
    Min: Integer;
    class operator GreaterThanOrEqual(a, b: TLimit): Boolean;
  end;
class operator TLimit.GreaterThanOrEqual(a, b: TLimit): Boolean;
begin Result := a.Min >= b.Min; end;
var a, b: TLimit;
begin
  a.Min := 0; b.Min := 0;
  if a >= b then WriteLn('ok') else WriteLn('under');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn overload_multiply_record_by_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPair = record
    V: Integer;
    class operator Multiply(a, b: TPair): TPair;
  end;
class operator TPair.Multiply(a, b: TPair): TPair;
begin Result.V := a.V * b.V; end;
var a, b, c: TPair;
begin
  a.V := 6; b.V := 7;
  c := a * b;
  WriteLn(c.V);
end."#
        ),
        &["42"]
    );
}

#[test]
fn overload_add_three_records_via_chaining() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAcc = record
    V: Integer;
    class operator Add(a, b: TAcc): TAcc;
  end;
class operator TAcc.Add(a, b: TAcc): TAcc;
begin Result.V := a.V + b.V; end;
var a, b, c, d: TAcc;
begin
  a.V := 1; b.V := 2; c.V := 3;
  d := a + b + c;
  WriteLn(d.V);
end."#
        ),
        &["6"]
    );
}

#[test]
fn overload_not_equal_on_strings_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TTag = record
    Name: String;
    class operator NotEqual(a, b: TTag): Boolean;
  end;
class operator TTag.NotEqual(a, b: TTag): Boolean;
begin Result := a.Name <> b.Name; end;
var a, b: TTag;
begin
  a.Name := 'alpha'; b.Name := 'beta';
  if a <> b then WriteLn('diff') else WriteLn('same');
end."#
        ),
        &["diff"]
    );
}

#[test]
fn overload_explicit_to_boolean_from_flag() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFlag = record
    On: Boolean;
    class operator Explicit(f: TFlag): Boolean;
  end;
class operator TFlag.Explicit(f: TFlag): Boolean;
begin Result := f.On; end;
var f: TFlag; b: Boolean;
begin
  f.On := true;
  b := Boolean(f);
  if b then WriteLn('on') else WriteLn('off');
end."#
        ),
        &["on"]
    );
}

#[test]
fn overload_inc_amount_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TTick = record
    N: Integer;
    class operator Inc(var t: TTick; k: Integer);
  end;
class operator TTick.Inc(var t: TTick; k: Integer);
begin t.N := t.N + k; end;
var t: TTick; i: Integer;
begin
  t.N := 0;
  for i := 1 to 3 do Inc(t, 2);
  WriteLn(t.N);
end."#
        ),
        &["6"]
    );
}

#[test]
fn overload_dec_amount_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TTick = record
    N: Integer;
    class operator Dec(var t: TTick; k: Integer);
  end;
class operator TTick.Dec(var t: TTick; k: Integer);
begin t.N := t.N - k; end;
var t: TTick; i: Integer;
begin
  t.N := 20;
  for i := 1 to 2 do Dec(t, 3);
  WriteLn(t.N);
end."#
        ),
        &["14"]
    );
}

#[test]
fn overload_compare_equal_false_for_different() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TId = record
    Value: Integer;
    class operator Equal(a, b: TId): Boolean;
  end;
class operator TId.Equal(a, b: TId): Boolean;
begin Result := a.Value = b.Value; end;
var a, b: TId;
begin
  a.Value := 1; b.Value := 2;
  if a = b then WriteLn('eq') else WriteLn('ne');
end."#
        ),
        &["ne"]
    );
}

#[test]
fn overload_less_than_false_when_greater() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TNum = record
    V: Integer;
    class operator LessThan(a, b: TNum): Boolean;
  end;
class operator TNum.LessThan(a, b: TNum): Boolean;
begin Result := a.V < b.V; end;
var a, b: TNum;
begin
  a.V := 50; b.V := 10;
  if a < b then WriteLn('yes') else WriteLn('no');
end."#
        ),
        &["no"]
    );
}

#[test]
fn overload_div_mod_combined_check() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TDivMod = record
    N: Integer;
    class operator IntDivide(d: TDivMod; k: Integer): Integer;
    class operator Modulus(d: TDivMod; k: Integer): Integer;
  end;
class operator TDivMod.IntDivide(d: TDivMod; k: Integer): Integer;
begin Result := d.N div k; end;
class operator TDivMod.Modulus(d: TDivMod; k: Integer): Integer;
begin Result := d.N mod k; end;
var d: TDivMod; q, r: Integer;
begin
  d.N := 53;
  q := d div 6;
  r := d mod 6;
  WriteLn(q); WriteLn(r);
end."#
        ),
        &["8", "5"]
    );
}
