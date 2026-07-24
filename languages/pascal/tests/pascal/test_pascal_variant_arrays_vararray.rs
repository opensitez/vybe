use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 93: Variant Arrays & Dynamic VarArray Manipulation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vararray_create_1d_integer() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 2], varInteger);
  v[0] := 10; v[1] := 20; v[2] := 30;
  WriteLn(v[0] + v[1] + v[2]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_vararray_of_inline_creation() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayOf(['Alpha', 'Beta', 'Gamma']);
  WriteLn(v[0]);
  WriteLn(v[1]);
  WriteLn(v[2]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["Alpha", "Beta", "Gamma"]);
}

#[test]
fn test_vararray_low_high_bound_dimcount() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([1, 5], varInteger);
  WriteLn(VarArrayDimCount(v));
  WriteLn(VarArrayLowBound(v, 1));
  WriteLn(VarArrayHighBound(v, 1));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["1", "1", "5"]);
}

#[test]
fn test_vararray_2d_create_and_access() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 1, 0, 1], varInteger);
  v[0, 0] := 1; v[0, 1] := 2;
  v[1, 0] := 3; v[1, 1] := 4;
  WriteLn(v[0, 0].ToString + ',' + v[0, 1].ToString);
  WriteLn(v[1, 0].ToString + ',' + v[1, 1].ToString);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["1,2", "3,4"]);
}

#[test]
fn test_vararray_2d_dimcount_bounds() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 2, 0, 4], varDouble);
  WriteLn(VarArrayDimCount(v));
  WriteLn(VarArrayHighBound(v, 1));
  WriteLn(VarArrayHighBound(v, 2));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["2", "2", "4"]);
}

#[test]
fn test_vararray_redim_resize() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 1], varInteger);
  v[0] := 100; v[1] := 200;
  VarArrayRedim(v, 3);
  v[2] := 300;
  WriteLn(v[0]);
  WriteLn(v[2]);
  WriteLn(VarArrayHighBound(v, 1));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["100", "300", "3"]);
}

#[test]
fn test_vararray_lock_unlock_pointer() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant; p: PInteger;
begin
  v := VarArrayCreate([0, 1], varInteger);
  v[0] := 77; v[1] := 88;
  p := VarArrayLock(v);
  try
    WriteLn(p^);
    WriteLn((p + 1)^);
  finally
    VarArrayUnlock(v);
  end;
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["77", "88"]);
}

#[test]
fn test_vararray_heterogeneous_varvariant_elements() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 2], varVariant);
  v[0] := 100;
  v[1] := 'TextValue';
  v[2] := 3.14;
  WriteLn(v[0]);
  WriteLn(v[1]);
  WriteLn(v[2]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["100", "TextValue", "3.14"]);
}

#[test]
fn test_vararray_varisarray_check() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := VarArrayCreate([0, 1], varInteger);
  v2 := 100;
  WriteLn(VarIsArray(v1));
  WriteLn(VarIsArray(v2));
  VarClear(v1);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_vararray_copy_assignment() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := VarArrayOf([10, 20]);
  v2 := v1;
  WriteLn(v2[0]);
  WriteLn(v2[1]);
  VarClear(v1); VarClear(v2);
end.
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_vararray_3d_cube_access() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 1, 0, 1, 0, 1], varInteger);
  v[1, 1, 1] := 999;
  WriteLn(v[1, 1, 1]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_vararray_element_mutation() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayOf([5, 10]);
  v[0] := v[0] + 15;
  WriteLn(v[0]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_vararray_loop_iteration() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant; i, sum: Integer;
begin
  v := VarArrayOf([1, 2, 3, 4]);
  sum := 0;
  for i := VarArrayLowBound(v, 1) to VarArrayHighBound(v, 1) do
    sum := sum + Integer(v[i]);
  WriteLn(sum);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_vararray_boolean_elements() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 1], varBoolean);
  v[0] := True; v[1] := False;
  WriteLn(v[0]);
  WriteLn(v[1]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_vararray_float_elements_sum() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayOf([1.5, 2.5]);
  WriteLn(Double(v[0]) + Double(v[1]));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_vararray_out_of_bounds_exception() {
    let out = run_pascal(r#"
program Test;
uses Variants, SysUtils;
var v: Variant; dummy: Variant;
begin
  v := VarArrayCreate([0, 1], varInteger);
  try
    dummy := v[5];
  except
    on E: EVariantArrayBoundsError do WriteLn('ArrayBoundsErrorCaught');
  end;
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["ArrayBoundsErrorCaught"]);
}

#[test]
fn test_vararray_string_search() {
    let out = run_pascal(r#"
program Test;
uses Variants;
function FindInVarArray(const arr: Variant; const target: String): Boolean;
var i: Integer;
begin
  Result := False;
  for i := VarArrayLowBound(arr, 1) to VarArrayHighBound(arr, 1) do
    if VarToStr(arr[i]) = target then Exit(True);
end;

var v: Variant;
begin
  v := VarArrayOf(['Apple', 'Banana', 'Cherry']);
  WriteLn(FindInVarArray(v, 'Banana'));
  WriteLn(FindInVarArray(v, 'Durian'));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_vararray_empty_array_check() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, -1], varInteger);
  WriteLn(VarArrayLowBound(v, 1) > VarArrayHighBound(v, 1));
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_vararray_byte_array_lock() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant; pb: PByte;
begin
  v := VarArrayCreate([0, 1], varByte);
  v[0] := $AA; v[1] := $BB;
  pb := VarArrayLock(v);
  try
    WriteLn(HexStr(pb^, 2));
    WriteLn(HexStr((pb + 1)^, 2));
  finally
    VarArrayUnlock(v);
  end;
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["AA", "BB"]);
}

#[test]
fn test_vararray_reassignment_clears_previous() {
    let out = run_pascal(r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayOf([10]);
  v := VarArrayOf([20, 30]);
  WriteLn(VarArrayHighBound(v, 1));
  WriteLn(v[1]);
  VarClear(v);
end.
"#);
    assert_eq!(out, vec!["1", "30"]);
}
