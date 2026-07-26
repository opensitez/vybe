use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 21: Pointer Type Casting & Address Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_pointer_basic_address_and_dereference() {
    let out = run_pascal(
        r#"
program Test;
var x: Integer;
    p: PInteger;
begin
  x := 42;
  p := @x;
  WriteLn(p^);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_untyped_pointer_assignment() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer;
    untypedPtr: Pointer;
    typedPtr: PInteger;
begin
  val := 100;
  untypedPtr := @val;
  typedPtr := PInteger(untypedPtr);
  WriteLn(typedPtr^);
end.
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_pchar_to_pbyte_casting() {
    let out = run_pascal(
        r#"
program Test;
var c: Char;
    pc: PChar;
    pb: PByte;
begin
  c := 'A';
  pc := @c;
  pb := PByte(pc);
  WriteLn(pb^);
end.
"#,
    );
    assert_eq!(out, vec!["65"]);
}

#[test]
fn test_double_pointer_dereference() {
    let out = run_pascal(
        r#"
program Test;
type PPInteger = ^PInteger;
var val: Integer;
    p1: PInteger;
    p2: PPInteger;
begin
  val := 777;
  p1 := @val;
  p2 := @p1;
  WriteLn(p2^^);
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_array_element_address() {
    let out = run_pascal(
        r#"
program Test;
var arr: array[1..3] of Integer;
    p: PInteger;
begin
  arr[1] := 10; arr[2] := 20; arr[3] := 30;
  p := @arr[2];
  WriteLn(p^);
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_record_field_address() {
    let out = run_pascal(
        r#"
program Test;
type TData = record Code: Integer; Name: String; end;
var rec: TData;
    pCode: PInteger;
begin
  rec.Code := 999; rec.Name := 'Test';
  pCode := @rec.Code;
  WriteLn(pCode^);
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_pointer_mutation_via_dereference() {
    let out = run_pascal(
        r#"
program Test;
var num: Integer;
    p: PInteger;
begin
  num := 10;
  p := @num;
  p^ := 50;
  WriteLn(num);
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_pointer_to_nativeint_casting() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer;
    p: PInteger;
    addr: NativeInt;
begin
  val := 123;
  p := @val;
  addr := NativeInt(p);
  WriteLn(addr > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pboolean_pointer() {
    let out = run_pascal(
        r#"
program Test;
var flag: Boolean;
    pb: PBoolean;
begin
  flag := True;
  pb := @flag;
  WriteLn(pb^);
  pb^ := False;
  WriteLn(flag);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_preal_pointer() {
    let out = run_pascal(
        r#"
program Test;
var num: Real;
    pr: PReal;
begin
  num := 3.14;
  pr := @num;
  WriteLn(pr^);
  pr^ := 2.71;
  WriteLn(num);
end.
"#,
    );
    assert_eq!(out, vec!["3.14", "2.71"]);
}

#[test]
fn test_pointer_comparisons() {
    let out = run_pascal(
        r#"
program Test;
var a, b: Integer;
    pa, pb, pa2: PInteger;
begin
  pa := @a; pb := @b; pa2 := @a;
  WriteLn(pa = pa2);
  WriteLn(pa <> pb);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_nil_pointer_casting() {
    let out = run_pascal(
        r#"
program Test;
var pUntyped: Pointer;
    pTyped: PInteger;
begin
  pUntyped := nil;
  pTyped := PInteger(pUntyped);
  WriteLn(pTyped = nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pword_pointer_typecast() {
    let out = run_pascal(
        r#"
program Test;
var w: Word;
    pw: PWord;
begin
  w := 65535;
  pw := @w;
  WriteLn(pw^);
end.
"#,
    );
    assert_eq!(out, vec!["65535"]);
}

#[test]
fn test_pointer_cast_inside_procedure() {
    let out = run_pascal(
        r#"
program Test;
procedure PrintIntPtr(ptr: Pointer);
var p: PInteger;
begin
  p := PInteger(ptr);
  WriteLn(p^);
end;
var x: Integer;
begin
  x := 888;
  PrintIntPtr(@x);
end.
"#,
    );
    assert_eq!(out, vec!["888"]);
}

#[test]
fn test_pointer_cast_function_return() {
    let out = run_pascal(
        r#"
program Test;
var globalVal: Integer;
function GetValPointer: Pointer;
begin
  globalVal := 555;
  Result := @globalVal;
end;
var p: PInteger;
begin
  p := PInteger(GetValPointer);
  WriteLn(p^);
end.
"#,
    );
    assert_eq!(out, vec!["555"]);
}

#[test]
fn test_pointer_arithmetic_dereference_expression() {
    let out = run_pascal(
        r#"
program Test;
var a, b: Integer;
    pa, pb: PInteger;
begin
  a := 10; b := 20;
  pa := @a; pb := @b;
  WriteLn(pa^ + pb^);
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_penum_pointer_typecast() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
type PColor = ^TColor;
var col: TColor;
    pc: PColor;
begin
  col := cBlue;
  pc := @col;
  WriteLn(Ord(pc^));
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_psubrange_pointer() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 1..100;
type PScore = ^TScore;
var sc: TScore;
    ps: PScore;
begin
  sc := 95;
  ps := @sc;
  WriteLn(ps^);
end.
"#,
    );
    assert_eq!(out, vec!["95"]);
}

#[test]
fn test_pointer_to_record_structure() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type PPoint = ^TPoint;
var pt: TPoint;
    p: PPoint;
begin
  pt.X := 15; pt.Y := 30;
  p := @pt;
  WriteLn(p^.X);
  WriteLn(p^.Y);
end.
"#,
    );
    assert_eq!(out, vec!["15", "30"]);
}

#[test]
fn test_pointer_cast_and_bitwise_operations() {
    let out = run_pascal(
        r#"
program Test;
var val: Integer;
    p: PInteger;
begin
  val := $F0F0;
  p := @val;
  p^ := p^ or $0F0F;
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["65535"]);
}
