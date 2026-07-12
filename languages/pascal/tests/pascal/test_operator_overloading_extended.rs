/// Additional operator overloads on records: subtract, multiply, compare, negate.
use super::helpers::run_pascal;

#[test]
fn operator_overload_subtract_vectors() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TVec = record
    X, Y: Integer;
    class operator Subtract(a, b: TVec): TVec;
  end;
class operator TVec.Subtract(a, b: TVec): TVec;
begin Result.X := a.X - b.X; Result.Y := a.Y - b.Y; end;
var a, b, c: TVec;
begin
  a.X := 10; a.Y := 8;
  b.X := 3; b.Y := 5;
  c := a - b;
  WriteLn(c.X); WriteLn(c.Y);
end."#
        ),
        &["7", "3"]
    );
}

#[test]
fn operator_overload_multiply_record_by_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TScale = record
    V: Integer;
    class operator Multiply(a: TScale; k: Integer): TScale;
  end;
class operator TScale.Multiply(a: TScale; k: Integer): TScale;
begin Result.V := a.V * k; end;
var s, t: TScale;
begin
  s.V := 6;
  t := s * 7;
  WriteLn(t.V);
end."#
        ),
        &["42"]
    );
}

#[test]
fn operator_overload_not_inverts_flag_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TFlag = record
    On: Boolean;
    class operator Not(f: TFlag): TFlag;
  end;
class operator TFlag.Not(f: TFlag): TFlag;
begin Result.On := not f.On; end;
var a, b: TFlag;
begin
  a.On := true;
  b := not a;
  if b.On then WriteLn('on') else WriteLn('off');
end."#
        ),
        &["off"]
    );
}

#[test]
fn operator_overload_greater_than_orders_points() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X: Integer;
    class operator GreaterThan(a, b: TPoint): Boolean;
  end;
class operator TPoint.GreaterThan(a, b: TPoint): Boolean;
begin Result := a.X > b.X; end;
var p, q: TPoint;
begin
  p.X := 5; q.X := 2;
  if p > q then WriteLn('gt') else WriteLn('le');
end."#
        ),
        &["gt"]
    );
}

#[test]
fn operator_overload_inc_increments_wrapped_counter() {
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
  c.N := 9;
  Inc(c);
  WriteLn(c.N);
end."#
        ),
        &["10"]
    );
}

#[test]
fn operator_overload_dec_decrements_wrapped_counter() {
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
  c.N := 4;
  Dec(c);
  WriteLn(c.N);
end."#
        ),
        &["3"]
    );
}

#[test]
fn operator_overload_explicit_to_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TWrap = record
    V: Integer;
    class operator Explicit(w: TWrap): Integer;
  end;
class operator TWrap.Explicit(w: TWrap): Integer;
begin Result := w.V; end;
var w: TWrap; n: Integer;
begin
  w.V := 55;
  n := Integer(w);
  WriteLn(n);
end."#
        ),
        &["55"]
    );
}

#[test]
fn operator_overload_concatenate_two_strings_via_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TText = record
    S: String;
    class operator Add(a, b: TText): TText;
  end;
class operator TText.Add(a, b: TText): TText;
begin Result.S := a.S + b.S; end;
var x, y, z: TText;
begin
  x.S := 'foo'; y.S := 'bar';
  z := x + y;
  WriteLn(z.S);
end."#
        ),
        &["foobar"]
    );
}
