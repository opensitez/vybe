use super::helpers::run_pascal;

#[test]
fn pointer_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  WriteLn(p^);
end."#
        ),
        &["10"]
    );
}

#[test]
fn pointer_modify() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  p^ := 20;
  WriteLn(x);
end."#
        ),
        &["20"]
    );
}

#[test]
fn pointer_return_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function MakePtr: ^Integer;
var
  x: Integer;
begin
  x := 42;
  Result := @x;
end;

var
  p: ^Integer;
begin
  p := MakePtr();
  WriteLn(p^);
end."#
        ),
        &["42"]
    );
}

#[test]
fn pointer_pass_to_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure SetValue(p: ^Integer; value: Integer);
begin
  p^ := value;
end;

var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  SetValue(p, 33);
  WriteLn(x);
end."#
        ),
        &["33"]
    );
}

#[test]
fn pointer_to_record_field_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;

var
  point: TPoint;
  p: ^TPoint;
begin
  point.X := 7;
  point.Y := 9;
  p := @point;
  p^.X := 11;
  WriteLn(p^.X + p^.Y);
end."#
        ),
        &["20"]
    );
}

#[test]
fn pointer_to_array_index_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TIntArray = array[0..2] of Integer;

var
  values: TIntArray;
  p: ^TIntArray;
begin
  values[0] := 1;
  values[1] := 2;
  values[2] := 3;
  p := @values;
  p^[1] := 10;
  WriteLn(values[0] + values[1] + values[2]);
end."#
        ),
        &["14"]
    );
}

#[test]
fn pointer_to_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  p: ^Integer;
  pp: ^^Integer;
begin
  x := 12;
  p := @x;
  pp := @p;
  pp^^ := 27;
  WriteLn(x);
end."#
        ),
        &["27"]
    );
}

#[test]
fn pointer_equality() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  y: Integer;
  p1: ^Integer;
  p2: ^Integer;
  p3: ^Integer;
begin
  x := 10;
  y := 20;
  p1 := @x;
  p2 := @x;
  p3 := @y;
  WriteLn(p1 = p2);
  WriteLn(p1 = p3);
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn pointer_inequality() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  y: Integer;
  p1: ^Integer;
  p2: ^Integer;
begin
  x := 10;
  y := 10;
  p1 := @x;
  p2 := @y;
  WriteLn(p1 <> p2);
end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_to_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  s: String;
  p: ^String;
begin
  s := 'hello';
  p := @s;
  p^ := 'world';
  WriteLn(s);
end."#
        ),
        &["world"]
    );
}

#[test]
fn pointer_to_boolean() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  b: Boolean;
  p: ^Boolean;
begin
  b := false;
  p := @b;
  p^ := true;
  WriteLn(b);
end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_nil_compare() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  p1: ^Integer;
  p2: ^Integer;
begin
  x := 10;
  p1 := nil;
  p2 := @x;
  WriteLn(p1 = nil);
  WriteLn(p2 = nil);
end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn pointer_nil_reassignment() {
    assert_eq!(
        run_pascal(
            r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  p := nil;
  WriteLn(p = nil);
end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_passthrough_call_layers() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Identity(p: ^Integer): ^Integer;
begin
  Result := p;
end;

function Bounce(p: ^Integer): ^Integer;
begin
  Result := Identity(p);
end;

var
  x: Integer;
  p: ^Integer;
begin
  x := 41;
  p := Bounce(@x);
  p^ := 42;
  WriteLn(x);
end."#
        ),
        &["42"]
    );
}

// -------------------------------------------------------------------
// from test_pointers_address_deref.rs
// -------------------------------------------------------------------
#[test]
fn address_of_integer_variable() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
    p: ^Integer;
begin
  x := 42;
  p := @x;
  WriteLn(p^);
end."#
        ),
        &["42"]
    );
}

#[test]
fn dereference_assign_updates_target() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
    p: ^Integer;
begin
  x := 1;
  p := @x;
  p^ := 9;
  WriteLn(x);
end."#
        ),
        &["9"]
    );
}

#[test]
fn pointer_to_char_writes_through_deref() {
    assert_eq!(
        run_pascal(
            r#"program T;
var c: Char;
    p: ^Char;
begin
  c := 'a';
  p := @c;
  p^ := 'Z';
  WriteLn(c);
end."#
        ),
        &["Z"]
    );
}

#[test]
fn pointer_nil_compare_before_assign() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
begin
  p := nil;
  if p = nil then WriteLn('nil') else WriteLn('set');
end."#
        ),
        &["nil"]
    );
}

#[test]
fn pointer_to_record_field_via_address() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPoint = record X, Y: Integer; end;
var pt: TPoint;
    px: ^Integer;
begin
  pt.X := 3;
  px := @pt.X;
  px^ := 8;
  WriteLn(pt.X);
end."#
        ),
        &["8"]
    );
}

#[test]
fn new_dispose_integer_heap_cell() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
begin
  New(p);
  p^ := 77;
  WriteLn(p^);
  Dispose(p);
end."#
        ),
        &["77"]
    );
}

#[test]
fn pointer_increment_through_array_index() {
    assert_eq!(
        run_pascal(
            r#"program T;
var arr: array[0..2] of Integer;
    p: ^Integer;
begin
  arr[0] := 10;
  arr[1] := 20;
  arr[2] := 30;
  p := @arr[0];
  WriteLn(p^);
  Inc(p);
  WriteLn(p^);
end."#
        ),
        &["10", "20"]
    );
}

#[test]
fn pointer_passed_to_procedure_by_value() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure WriteThrough(p: ^Integer);
begin
  WriteLn(p^);
end;
var x: Integer;
begin
  x := 55;
  WriteThrough(@x);
end."#
        ),
        &["55"]
    );
}

#[test]
fn assigned_false_for_nil_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
begin
  p := nil;
  WriteLn(Assigned(p));
end."#
        ),
        &["false"]
    );
}

#[test]
fn assigned_true_after_address_taken() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
    p: ^Integer;
begin
  p := @x;
  WriteLn(Assigned(p));
end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_comparison_same_variable_address() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
    p, q: ^Integer;
begin
  p := @x;
  q := @x;
  WriteLn(p = q);
end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_to_pointer_deref_twice() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
    p: ^Integer;
    pp: ^^Integer;
begin
  x := 21;
  p := @x;
  pp := @p;
  WriteLn(pp^^);
end."#
        ),
        &["21"]
    );
}

#[test]
fn getmem_freemem_manual_lifetime() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: Pointer;
begin
  GetMem(p, SizeOf(Integer));
  PInteger(p)^ := 88;
  WriteLn(PInteger(p)^);
  FreeMem(p);
end."#
        ),
        &["88"]
    );
}

#[test]
fn typed_pointer_subscript_on_heap_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
var p: ^Integer;
    i: Integer;
begin
  New(p);
  for i := 0 to 2 do
    (p + i)^ := i + 1;
  WriteLn((p + 1)^);
  Dispose(p);
end."#
        ),
        &["2"]
    );
}

#[test]
fn assigned_returns_false_for_nil_pointer() {
    assert_eq!(
        run_pascal(r#"program T; var p: ^Integer; begin WriteLn(Assigned(p)); end."#),
        &["false"]
    );
}

#[test]
fn assigned_returns_true_after_new() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: ^Integer; begin New(p); WriteLn(Assigned(p)); Dispose(p); end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_equality_same_address() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; p, q: ^Integer; begin x := 1; p := @x; q := @x; WriteLn(p = q); end."#
        ),
        &["true"]
    );
}

#[test]
fn nil_pointer_compare_not_equal_assigned() {
    assert_eq!(
        run_pascal(r#"program T; var p: ^Integer; begin WriteLn(p = nil); end."#),
        &["true"]
    );
}

#[test]
fn dispose_sets_pointer_invalid_after_free() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: ^Integer; begin New(p); p^ := 3; Dispose(p); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}
