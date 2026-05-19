use super::helpers::run_pascal;

#[test] fn pointer_basic() {
    assert_eq!(run_pascal(r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  WriteLn(p^);
end."#), &["10"]);
}

#[test] fn pointer_modify() {
    assert_eq!(run_pascal(r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  p^ := 20;
  WriteLn(x);
end."#), &["20"]);
}

#[test] fn pointer_return_from_function() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["42"]);
}

#[test] fn pointer_pass_to_procedure() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["33"]);
}

#[test] fn pointer_to_record_field_access() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["20"]);
}

#[test] fn pointer_to_array_index_access() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["14"]);
}

#[test] fn pointer_to_pointer() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["27"]);
}

#[test] fn pointer_equality() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["true", "false"]);
}

#[test] fn pointer_inequality() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["true"]);
}

#[test] fn pointer_to_string() {
    assert_eq!(run_pascal(r#"program T;
var
  s: String;
  p: ^String;
begin
  s := 'hello';
  p := @s;
  p^ := 'world';
  WriteLn(s);
end."#), &["world"]);
}

#[test] fn pointer_to_boolean() {
    assert_eq!(run_pascal(r#"program T;
var
  b: Boolean;
  p: ^Boolean;
begin
  b := false;
  p := @b;
  p^ := true;
  WriteLn(b);
end."#), &["true"]);
}

#[test] fn pointer_nil_compare() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["true", "false"]);
}

#[test] fn pointer_nil_reassignment() {
    assert_eq!(run_pascal(r#"program T;
var
  x: Integer;
  p: ^Integer;
begin
  x := 10;
  p := @x;
  p := nil;
  WriteLn(p = nil);
end."#), &["true"]);
}

#[test] fn pointer_passthrough_call_layers() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["42"]);
}