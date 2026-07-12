/// Nested records, classes, and procedures — type nesting patterns.
use super::helpers::run_pascal;

#[test]
fn nested_record_inside_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TInner = record V: Integer; end;
  TOuter = record Inner: TInner; end;
var o: TOuter;
begin
  o.Inner.V := 12;
  WriteLn(o.Inner.V);
end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_record_three_levels_deep() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TL3 = record Z: Integer; end;
  TL2 = record Child: TL3; end;
  TL1 = record Child: TL2; end;
var root: TL1;
begin
  root.Child.Child.Z := 99;
  WriteLn(root.Child.Child.Z);
end."#
        ),
        &["99"]
    );
}

#[test]
fn nested_class_inside_program_scope() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TOuter = class
  type
    TInner = class
      V: Integer;
    end;
  public
    function Make: TInner;
  end;
function TOuter.Make: TInner;
begin Result := TInner.Create; Result.V := 7; end;
var o: TOuter; i: TOuter.TInner;
begin
  o := TOuter.Create;
  i := o.Make;
  WriteLn(i.V);
  i.Free; o.Free;
end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_procedure_inside_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
  procedure Inner;
  begin WriteLn('inner'); end;
begin Inner; end;
begin Outer; end."#
        ),
        &["inner"]
    );
}

#[test]
fn nested_function_inside_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Outer: Integer;
  function Inner: Integer;
  begin Result := 21; end;
begin Result := Inner * 2; end;
begin WriteLn(Outer); end."#
        ),
        &["42"]
    );
}

#[test]
fn local_type_alias_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Demo;
type TLocal = Integer;
var x: TLocal;
begin x := 8; WriteLn(x); end;
begin Demo; end."#
        ),
        &["8"]
    );
}

#[test]
fn nested_array_of_nested_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCell = record V: Integer; end;
var grid: array[0..1] of TCell;
    i: Integer; s: Integer;
begin
  grid[0].V := 2; grid[1].V := 3;
  s := 0;
  for i := 0 to 1 do s := s + grid[i].V;
  WriteLn(s);
end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_class_field_access_chain() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TNode = class
  Value: Integer;
  Next: TNode;
end;
var a, b: TNode;
begin
  a := TNode.Create; b := TNode.Create;
  a.Value := 1; b.Value := 2;
  a.Next := b;
  WriteLn(a.Next.Value);
  b.Free; a.Free;
end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_record_method_on_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TInner = record V: Integer; end;
  TOuter = record
    Inner: TInner;
    function Sum: Integer;
  end;
function TOuter.Sum: Integer;
begin Result := Inner.V + 1; end;
var o: TOuter;
begin
  o.Inner.V := 10;
  WriteLn(o.Sum);
end."#
        ),
        &["11"]
    );
}

#[test]
fn nested_enum_inside_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TState = (Off, On);
  TDevice = record Mode: TState; end;
var d: TDevice;
begin
  d.Mode := On;
  WriteLn(Ord(d.Mode));
end."#
        ),
        &["1"]
    );
}

#[test]
fn nested_set_field_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TD = (A, B, C);
type TBag = record Items: set of TD; end;
var bag: TBag;
begin
  bag.Items := [A, C];
  if B in bag.Items then WriteLn('has') else WriteLn('miss');
end."#
        ),
        &["miss"]
    );
}

#[test]
fn nested_procedure_in_class() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TWorker = class
  procedure Run;
end;
procedure TWorker.Run;
  procedure Step;
  begin WriteLn('step'); end;
begin Step; end;
var w: TWorker;
begin w := TWorker.Create; w.Run; w.Free; end."#
        ),
        &["step"]
    );
}

#[test]
fn nested_function_in_class_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCalc = class
  function Double(n: Integer): Integer;
end;
function TCalc.Double(n: Integer): Integer;
  function Inner(x: Integer): Integer;
  begin Result := x * 2; end;
begin Result := Inner(n); end;
var c: TCalc;
begin c := TCalc.Create; WriteLn(c.Double(6)); c.Free; end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_record_with_array_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBuf = record Data: array[0..2] of Integer; end;
var b: TBuf;
begin
  b.Data[1] := 44;
  WriteLn(b.Data[1]);
end."#
        ),
        &["44"]
    );
}

#[test]
fn nested_class_inherits_and_extends() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBase = class V: Integer; end;
type TDerived = class(TBase) W: Integer; end;
var d: TDerived;
begin
  d := TDerived.Create;
  d.V := 3; d.W := 4;
  WriteLn(d.V + d.W);
  d.Free;
end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_variant_record_case() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TKind = (IntK, StrK);
  TVar = record
    case K: TKind of
      IntK: (I: Integer);
      StrK: (S: String);
  end;
var v: TVar;
begin
  v.K := IntK;
  v.I := 55;
  WriteLn(v.I);
end."#
        ),
        &["55"]
    );
}

#[test]
fn nested_pointer_to_inner_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record V: Integer; end;
type TOuter = record P: ^TInner; end;
var inner: TInner; outer: TOuter;
begin
  inner.V := 18;
  outer.P := @inner;
  WriteLn(outer.P^.V);
end."#
        ),
        &["18"]
    );
}

#[test]
fn nested_type_in_case_statement() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCode = (A, B, C);
var c: TCode;
begin
  c := B;
  case c of
    A: WriteLn('a');
    B: WriteLn('b');
    C: WriteLn('c');
  end;
end."#
        ),
        &["b"]
    );
}

#[test]
fn nested_with_on_inner_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record X, Y: Integer; end;
type TOuter = record Inner: TInner; end;
var o: TOuter;
begin
  o.Inner.X := 2; o.Inner.Y := 3;
  with o.Inner do WriteLn(X + Y);
end."#
        ),
        &["5"]
    );
}

#[test]
fn nested_procedure_params_access_outer_type() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TItem = record V: Integer; end;
procedure Process(var item: TItem);
begin item.V := item.V * 2; end;
var it: TItem;
begin
  it.V := 5;
  Process(it);
  WriteLn(it.V);
end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_record_in_dynamic_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRow = record V: Integer; end;
var rows: array of TRow;
begin
  SetLength(rows, 2);
  rows[0].V := 4; rows[1].V := 6;
  WriteLn(rows[0].V + rows[1].V);
end."#
        ),
        &["10"]
    );
}

#[test]
fn nested_class_nested_type_reference() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TContainer = class
  type TItem = record N: Integer; end;
  Data: TItem;
end;
var c: TContainer;
begin
  c := TContainer.Create;
  c.Data.N := 13;
  WriteLn(c.Data.N);
  c.Free;
end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_function_returns_nested_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair = record A, B: Integer; end;
function Make: TPair;
begin Result.A := 1; Result.B := 2; end;
var p: TPair;
begin
  p := Make;
  WriteLn(p.A + p.B);
end."#
        ),
        &["3"]
    );
}

#[test]
fn nested_procedure_mutual_in_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
  procedure A; forward;
  procedure B;
  begin WriteLn('b'); A; end;
  procedure A;
  begin WriteLn('a'); end;
begin B; end;
begin Outer; end."#
        ),
        &["b", "a"]
    );
}

#[test]
fn nested_record_copy_inner_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record V: Integer; end;
type TOuter = record Inner: TInner; end;
var a, b: TOuter;
begin
  a.Inner.V := 7;
  b := a;
  b.Inner.V := 9;
  WriteLn(a.Inner.V); WriteLn(b.Inner.V);
end."#
        ),
        &["7", "9"]
    );
}

#[test]
fn nested_class_property_wraps_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TWrap = class
private FV: Integer;
public property Value: Integer read FV write FV;
end;
var w: TWrap;
begin
  w := TWrap.Create;
  w.Value := 25;
  WriteLn(w.Value);
  w.Free;
end."#
        ),
        &["25"]
    );
}

#[test]
fn nested_interface_style_method_on_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCounter = record
  N: Integer;
  procedure IncN;
end;
procedure TCounter.IncN;
begin N := N + 1; end;
var c: TCounter;
begin
  c.N := 0;
  c.IncN; c.IncN;
  WriteLn(c.N);
end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_subrange_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDigit = 0..9;
type TCell = record D: TDigit; end;
var c: TCell;
begin
  c.D := 7;
  WriteLn(c.D);
end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_static_array_of_nested_class_refs() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TItem = class V: Integer; end;
var items: array[0..1] of TItem;
begin
  items[0] := TItem.Create; items[1] := TItem.Create;
  items[0].V := 2; items[1].V := 5;
  WriteLn(items[0].V + items[1].V);
  items[1].Free; items[0].Free;
end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_record_in_function_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPt = record X, Y: Integer; end;
function Area(const p: TPt): Integer;
begin Result := p.X * p.Y; end;
var p: TPt;
begin
  p.X := 3; p.Y := 4;
  WriteLn(Area(p));
end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_type_forward_declaration() {
    assert_eq!(
        run_pascal(
            r#"program T;
type PNode = ^TNode;
     TNode = record Next: PNode; V: Integer; end;
var a, b: TNode;
    pa, pb: PNode;
begin
  a.V := 1; b.V := 2;
  pa := @a; pb := @b;
  pa^.Next := pb;
  WriteLn(pa^.Next^.V);
end."#
        ),
        &["2"]
    );
}

#[test]
fn nested_procedure_local_var_shadows_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
procedure Outer;
var x: Integer;
  procedure Inner;
  begin x := 5; WriteLn(x); end;
begin
  x := 1;
  Inner;
  WriteLn(x);
end;
begin x := 0; Outer; WriteLn(x); end."#
        ),
        &["5", "1", "0"]
    );
}

#[test]
fn nested_class_constructor_initializes_nested_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record V: Integer; end;
type TOuter = class
  Inner: TInner;
  constructor Create;
end;
constructor TOuter.Create;
begin Inner.V := 42; end;
var o: TOuter;
begin
  o := TOuter.Create;
  WriteLn(o.Inner.V);
  o.Free;
end."#
        ),
        &["42"]
    );
}

#[test]
fn nested_record_array_multidim_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCell = record V: Integer; end;
var m: array[0..1, 0..1] of TCell;
begin
  m[0,1].V := 6; m[1,0].V := 7;
  WriteLn(m[0,1].V + m[1,0].V);
end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_function_type_in_record_scope() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMath = record
    function Apply(f: function(x: Integer): Integer; n: Integer): Integer;
  end;
function TMath.Apply(f: function(x: Integer): Integer; n: Integer): Integer;
begin Result := f(n); end;
function Double(x: Integer): Integer;
begin Result := x * 2; end;
var m: TMath;
begin WriteLn(m.Apply(@Double, 11)); end."#
        ),
        &["22"]
    );
}

#[test]
fn nested_anonymous_struct_style_aggregate() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPair = record A, B: Integer; end;
var p: TPair;
begin
  p := Default(TPair);
  p.A := 3; p.B := 4;
  WriteLn(p.A + p.B);
end."#
        ),
        &["7"]
    );
}

#[test]
fn nested_class_destructor_cleanup() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRes = class
  V: Integer;
  destructor Destroy; override;
end;
destructor TRes.Destroy;
begin WriteLn('gone'); inherited; end;
var r: TRes;
begin
  r := TRes.Create;
  r.V := 1;
  r.Free;
end."#
        ),
        &["gone"]
    );
}

#[test]
fn nested_record_string_field_nested_concat() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TName = record First, Last: String; end;
function Full(const n: TName): String;
begin Result := n.First + ' ' + n.Last; end;
var n: TName;
begin
  n.First := 'Ada'; n.Last := 'Lovelace';
  WriteLn(Full(n));
end."#
        ),
        &["Ada Lovelace"]
    );
}

#[test]
fn nested_procedure_in_nested_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Outer;
  procedure Mark(row, col: Integer);
  begin WriteLn(row * 10 + col); end;
  r, c: Integer;
begin
  for r := 0 to 1 do
    for c := 0 to 1 do
      Mark(r, c);
end;
begin Outer; end."#
        ),
        &["0", "1", "10", "11"]
    );
}

#[test]
fn nested_type_used_in_generic_style_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox = record Payload: Integer; end;
type TList = record
  Items: array[0..1] of TBox;
  function Get(i: Integer): Integer;
end;
function TList.Get(i: Integer): Integer;
begin Result := Items[i].Payload; end;
var L: TList;
begin
  L.Items[0].Payload := 8;
  L.Items[1].Payload := 9;
  WriteLn(L.Get(0) + L.Get(1));
end."#
        ),
        &["17"]
    );
}
