use super::helpers::run_pascal;

#[test]
fn test_variant_record_point() {
    let src = r#"
program T;
type
  TShape = record
    case Tag: Integer of
      0: (PX, PY: Integer);
      1: (Radius: Integer);
  end;
var
  s: TShape;
begin
  s.Tag := 0;
  s.PX := 3;
  s.PY := 4;
  WriteLn(s.Tag);
  WriteLn(s.PX);
  WriteLn(s.PY);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "3", "4"]);
}

#[test]
fn test_variant_record_circle() {
    let src = r#"
program T;
type
  TShape = record
    case Tag: Integer of
      0: (PX, PY: Integer);
      1: (Radius: Integer);
  end;
var
  s: TShape;
begin
  s.Tag := 1;
  s.Radius := 10;
  WriteLn(s.Tag);
  WriteLn(s.Radius);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "10"]);
}

#[test]
fn test_variant_record_case_dispatch() {
    let src = r#"
program T;
type
  TNum = record
    case IsReal: Boolean of
      false: (IVal: Integer);
      true:  (FVal: Double);
  end;
var
  n: TNum;
begin
  n.IsReal := false;
  n.IVal := 42;
  if not n.IsReal then
    WriteLn(n.IVal)
  else
    WriteLn(n.FVal);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_variant_record_with_fixed_fields() {
    let src = r#"
program T;
type
  TMsg = record
    ID: Integer;
    case Kind: Integer of
      1: (Text: string);
      2: (Code: Integer);
  end;
var
  m: TMsg;
begin
  m.ID := 99;
  m.Kind := 1;
  m.Text := 'hello';
  WriteLn(m.ID);
  WriteLn(m.Kind);
  WriteLn(m.Text);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["99", "1", "hello"]);
}

#[test]
fn test_variant_record_in_array() {
    let src = r#"
program T;
type
  TValue = record
    case T: Integer of
      0: (I: Integer);
      1: (S: string);
  end;
var
  arr: array[0..1] of TValue;
begin
  arr[0].T := 0; arr[0].I := 7;
  arr[1].T := 1; arr[1].S := 'hi';
  WriteLn(arr[0].I);
  WriteLn(arr[1].S);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["7", "hi"]);
}

#[test]
fn test_variant_record_reassign_tag() {
    let src = r#"
program T;
type
  TVar = record
    case Kind: Integer of
      0: (A, B: Integer);
      1: (C: Integer);
  end;
var
  v: TVar;
begin
  v.Kind := 0;
  v.A := 1;
  v.B := 2;
  WriteLn(v.A + v.B);
  v.Kind := 1;
  v.C := 99;
  WriteLn(v.C);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3", "99"]);
}

#[test]
fn test_variant_record_function_param() {
    let src = r#"
program T;
type
  TToken = record
    case Kind: Integer of
      0: (IntVal: Integer);
      1: (StrVal: string);
  end;

function TokenToStr(t: TToken): string;
begin
  case t.Kind of
    0: Result := 'int:' + IntToStr(t.IntVal);
    1: Result := 'str:' + t.StrVal;
    else Result := 'unknown';
  end;
end;

var
  t: TToken;
begin
  t.Kind := 0;
  t.IntVal := 42;
  WriteLn(TokenToStr(t));
  t.Kind := 1;
  t.StrVal := 'abc';
  WriteLn(TokenToStr(t));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["int:42", "str:abc"]);
}

#[test]
fn test_variant_record_boolean_discriminant() {
    let src = r#"
program T;
type
  TResult = record
    OK: Boolean;
    case Tag: Integer of
      0: (ErrMsg: string);
      1: (Value: Integer);
  end;
var
  r: TResult;
begin
  r.OK := true;
  r.Tag := 1;
  r.Value := 100;
  if r.OK then
    WriteLn('ok:' + IntToStr(r.Value))
  else
    WriteLn('err:' + r.ErrMsg);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["ok:100"]);
}

#[test]
fn test_variant_record_error_case() {
    let src = r#"
program T;
type
  TResult = record
    OK: Boolean;
    case Tag: Integer of
      0: (ErrMsg: string);
      1: (Value: Integer);
  end;
var
  r: TResult;
begin
  r.OK := false;
  r.Tag := 0;
  r.ErrMsg := 'not found';
  if r.OK then
    WriteLn('value')
  else
    WriteLn('err:' + r.ErrMsg);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["err:not found"]);
}

#[test]
fn test_variant_record_three_variants() {
    let src = r#"
program T;
type
  TShape = record
    case ShapeType: Integer of
      0: (X, Y: Integer);
      1: (Radius: Integer);
      2: (W, H: Integer);
  end;
procedure Describe(s: TShape);
begin
  case s.ShapeType of
    0: WriteLn('point(' + IntToStr(s.X) + ',' + IntToStr(s.Y) + ')');
    1: WriteLn('circle(r=' + IntToStr(s.Radius) + ')');
    2: WriteLn('rect(' + IntToStr(s.W) + 'x' + IntToStr(s.H) + ')');
  end;
end;
var
  p, c, r: TShape;
begin
  p.ShapeType := 0; p.X := 1; p.Y := 2;
  c.ShapeType := 1; c.Radius := 5;
  r.ShapeType := 2; r.W := 4; r.H := 3;
  Describe(p);
  Describe(c);
  Describe(r);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["point(1,2)", "circle(r=5)", "rect(4x3)"]);
}

#[test]
fn test_variant_record_count_by_tag() {
    let src = r#"
program T;
type
  TItem = record
    case Kind: Integer of
      0: (Count: Integer);
      1: (Name: string);
  end;
var
  items: array[0..3] of TItem;
  i, cnt: Integer;
begin
  items[0].Kind := 0; items[0].Count := 5;
  items[1].Kind := 1; items[1].Name := 'a';
  items[2].Kind := 0; items[2].Count := 3;
  items[3].Kind := 1; items[3].Name := 'b';
  cnt := 0;
  for i := 0 to 3 do
    if items[i].Kind = 0 then
      cnt := cnt + items[i].Count;
  WriteLn(cnt);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_variant_record_equality_check() {
    let src = r#"
program T;
type
  TPair = record
    case IsStr: Boolean of
      false: (N: Integer);
      true:  (S: string);
  end;
var
  a, b: TPair;
begin
  a.IsStr := false; a.N := 42;
  b.IsStr := false; b.N := 42;
  if (a.IsStr = b.IsStr) and (a.N = b.N) then
    WriteLn('equal')
  else
    WriteLn('different');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["equal"]);
}

#[test]
fn test_variant_record_nested_case() {
    let src = r#"
program T;
type
  TEvent = record
    Timestamp: Integer;
    case EventType: Integer of
      1: (KeyCode: Integer);
      2: (MouseX, MouseY: Integer);
      3: (WindowID: Integer);
  end;
var
  ev: TEvent;
begin
  ev.Timestamp := 100;
  ev.EventType := 2;
  ev.MouseX := 320;
  ev.MouseY := 240;
  WriteLn(ev.Timestamp);
  WriteLn(ev.EventType);
  WriteLn(ev.MouseX);
  WriteLn(ev.MouseY);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["100", "2", "320", "240"]);
}

#[test]
fn test_variant_record_union_memory() {
    let src = r#"
program T;
type
  TUnion = record
    case T: Integer of
      0: (Lo, Hi: Integer);
      1: (Full: Integer);
  end;
var
  u: TUnion;
begin
  u.T := 0;
  u.Lo := 1;
  u.Hi := 2;
  WriteLn(u.Lo);
  WriteLn(u.Hi);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_variant_record_serialize() {
    let src = r#"
program T;
type
  TVal = record
    case VType: Integer of
      0: (VInt: Integer);
      1: (VBool: Boolean);
      2: (VStr: string);
  end;

function Serialize(v: TVal): string;
begin
  case v.VType of
    0: Result := 'I:' + IntToStr(v.VInt);
    1: Result := 'B:' + BoolToStr(v.VBool);
    2: Result := 'S:' + v.VStr;
    else Result := '?';
  end;
end;

var
  vi, vb, vs: TVal;
begin
  vi.VType := 0; vi.VInt := 123;
  vb.VType := 1; vb.VBool := true;
  vs.VType := 2; vs.VStr := 'hello';
  WriteLn(Serialize(vi));
  WriteLn(Serialize(vb));
  WriteLn(Serialize(vs));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["I:123", "B:true", "S:hello"]);
}

#[test]
fn test_variant_record_swap_kinds() {
    let src = r#"
program T;
type
  TData = record
    case Kind: Integer of
      0: (IntData: Integer);
      1: (StrData: string);
  end;
procedure PrintData(d: TData);
begin
  if d.Kind = 0 then
    WriteLn(IntToStr(d.IntData))
  else
    WriteLn(d.StrData);
end;
var
  d: TData;
begin
  d.Kind := 0; d.IntData := 7;
  PrintData(d);
  d.Kind := 1; d.StrData := 'seven';
  PrintData(d);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["7", "seven"]);
}

#[test]
fn test_variant_record_in_function_return() {
    let src = r#"
program T;
type
  TBox = record
    case IsEmpty: Boolean of
      false: (Value: Integer);
      true: ();
  end;

function MakeBox(v: Integer): TBox;
begin
  Result.IsEmpty := false;
  Result.Value := v;
end;

function EmptyBox: TBox;
begin
  Result.IsEmpty := true;
end;

procedure Unbox(b: TBox);
begin
  if not b.IsEmpty then
    WriteLn(b.Value)
  else
    WriteLn('empty');
end;

begin
  Unbox(MakeBox(42));
  Unbox(EmptyBox);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42", "empty"]);
}

#[test]
fn test_variant_record_max_selector() {
    let src = r#"
program T;
type
  TNumber = record
    case IsFloat: Boolean of
      false: (IntN: Integer);
      true:  (FloatN: Double);
  end;
function GetInt(n: TNumber): Integer;
begin
  if not n.IsFloat then
    Result := n.IntN
  else
    Result := Round(n.FloatN);
end;
var
  a, b: TNumber;
begin
  a.IsFloat := false; a.IntN := 10;
  b.IsFloat := true; b.FloatN := 7.8;
  WriteLn(GetInt(a));
  WriteLn(GetInt(b));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["10", "8"]);
}

#[test]
fn test_variant_record_linked_kind() {
    let src = r#"
program T;
type
  TNode = record
    Next: Integer;
    case Leaf: Boolean of
      true:  (LeafVal: Integer);
      false: (ChildIdx: Integer);
  end;
var
  n: TNode;
begin
  n.Next := 2;
  n.Leaf := true;
  n.LeafVal := 99;
  WriteLn(n.Next);
  WriteLn(n.LeafVal);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["2", "99"]);
}

#[test]
fn test_variant_record_passthrough_array() {
    let src = r#"
program T;
type
  TItem = record
    ID: Integer;
    case IsPrimary: Boolean of
      true:  (Priority: Integer);
      false: (SecondaryID: Integer);
  end;
var
  items: array[0..1] of TItem;
  i: Integer;
begin
  items[0].ID := 1; items[0].IsPrimary := true; items[0].Priority := 10;
  items[1].ID := 2; items[1].IsPrimary := false; items[1].SecondaryID := 99;
  for i := 0 to 1 do begin
    Write(IntToStr(items[i].ID) + ':');
    if items[i].IsPrimary then
      WriteLn('p=' + IntToStr(items[i].Priority))
    else
      WriteLn('s=' + IntToStr(items[i].SecondaryID));
  end;
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1:p=10", "2:s=99"]);
}

#[test]
fn variant_record_boolean_tag_selects_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TVal = record
  case Boolean of
    True: (I: Integer);
    False: (S: String);
end;
var v: TVal;
begin
  v.I := 7;
  WriteLn(v.I);
end."#
        ),
        &["7"]
    );
}

#[test]
fn variant_record_integer_tag_two_shapes() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TNum = record
  case Byte of
    1: (A: Integer);
    2: (B, C: Integer);
end;
var n: TNum;
begin
  n.A := 5;
  WriteLn(n.A);
end."#
        ),
        &["5"]
    );
}

#[test]
fn variant_record_enum_discriminant() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TKind = (Circle, Rect);
    TShape = record
  case TKind of
    Circle: (Radius: Real);
    Rect: (W, H: Real);
end;
var s: TShape;
begin
  s.Radius := 3.0;
  WriteLn(Trunc(s.Radius));
end."#
        ),
        &["3"]
    );
}

#[test]
fn variant_record_with_shared_prefix_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TTagged = record
  Tag: Integer;
  case Integer of
    0: (N: Integer);
    1: (Text: String);
end;
var t: TTagged;
begin
  t.Tag := 0;
  t.N := 42;
  WriteLn(t.N);
end."#
        ),
        &["42"]
    );
}

#[test]
fn variant_record_assign_between_same_tag() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair = record
  case Boolean of
    True: (X: Integer);
    False: (Y: Integer);
end;
var a, b: TPair;
begin
  a.X := 10;
  b := a;
  WriteLn(b.X);
end."#
        ),
        &["10"]
    );
}
