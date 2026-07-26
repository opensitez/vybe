use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 5: Variant Types, Coercion & Dynamic Dispatch
// ═══════════════════════════════════════════════════════════

#[test]
fn test_variant_integer_and_string_assignment() {
    let out = run_pascal(
        r#"
program Test;
var v: Variant;
begin
  v := 42;
  WriteLn(v);
  v := 'Hello Variant';
  WriteLn(v);
end.
"#,
    );
    assert_eq!(out, vec!["42", "Hello Variant"]);
}

#[test]
fn test_variant_implicit_string_to_int_coercion() {
    let out = run_pascal(
        r#"
program Test;
var v1, v2: Variant;
    sum: Integer;
begin
  v1 := '100';
  v2 := '200';
  sum := v1 + v2;
  WriteLn(sum);
end.
"#,
    );
    assert_eq!(out, vec!["300"]);
}

#[test]
fn test_variant_unassigned_and_null_checks() {
    let out = run_pascal(
        r#"
program Test;
var v: Variant;
begin
  WriteLn(VarIsEmpty(v));
  v := Null;
  WriteLn(VarIsNull(v));
  WriteLn(VarIsEmpty(v));
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_variant_floating_point_arithmetic() {
    let out = run_pascal(
        r#"
program Test;
var v1, v2, res: Variant;
begin
  v1 := 12.5;
  v2 := 2.5;
  res := v1 * v2;
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["31.25"]);
}

#[test]
fn test_variant_boolean_logical_operations() {
    let out = run_pascal(
        r#"
program Test;
var b1, b2, res: Variant;
begin
  b1 := True;
  b2 := False;
  res := b1 and (not b2);
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_string_concatenation() {
    let out = run_pascal(
        r#"
program Test;
var s1, s2, res: Variant;
begin
  s1 := 'Object ';
  s2 := 'Pascal';
  res := s1 + s2;
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["Object Pascal"]);
}

#[test]
fn test_variant_array_creation_and_indexing() {
    let out = run_pascal(
        r#"
program Test;
var vArr: Variant;
begin
  vArr := VarArrayCreate([0, 2], varInteger);
  vArr[0] := 10;
  vArr[1] := 20;
  vArr[2] := 30;
  WriteLn(vArr[0]);
  WriteLn(vArr[1]);
  WriteLn(vArr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_variant_array_bounds_queries() {
    let out = run_pascal(
        r#"
program Test;
var vArr: Variant;
begin
  vArr := VarArrayCreate([1, 5], varString);
  WriteLn(VarArrayLowBound(vArr, 1));
  WriteLn(VarArrayHighBound(vArr, 1));
end.
"#,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn test_variant_type_casting_varastype() {
    let out = run_pascal(
        r#"
program Test;
var v: Variant;
    strVal: String;
begin
  v := 999;
  v := VarAsType(v, varString);
  strVal := v;
  WriteLn(strVal);
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_variant_vartype_code_inspection() {
    let out = run_pascal(
        r#"
program Test;
var v: Variant;
begin
  v := 100;
  WriteLn(VarType(v) = varInteger);
  v := 'text';
  WriteLn(VarType(v) = varString);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_variant_comparison_operators() {
    let out = run_pascal(
        r#"
program Test;
var v1, v2: Variant;
begin
  v1 := 50;
  v2 := 100;
  WriteLn(v1 < v2);
  WriteLn(v1 = 50);
  WriteLn(v2 <> 50);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_variant_parameter_passing() {
    let out = run_pascal(
        r#"
program Test;
function FormatVariant(v: Variant): String;
begin
  Result := 'Val=' + VarToStr(v);
end;
begin
  WriteLn(FormatVariant(123));
  WriteLn(FormatVariant('ABC'));
end.
"#,
    );
    assert_eq!(out, vec!["Val=123", "Val=ABC"]);
}

#[test]
fn test_variant_function_return_variant() {
    let out = run_pascal(
        r#"
program Test;
function MultiplyVariants(a, b: Variant): Variant;
begin
  Result := a * b;
end;
begin
  WriteLn(MultiplyVariants(6, 7));
  WriteLn(MultiplyVariants(1.5, 4.0));
end.
"#,
    );
    assert_eq!(out, vec!["42", "6"]);
}

#[test]
fn test_variant_record_field_storage() {
    let out = run_pascal(
        r#"
program Test;
type TDataField = record
  FieldName: String;
  Value: Variant;
end;
var field: TDataField;
begin
  field.FieldName := 'Price';
  field.Value := 19.99;
  WriteLn(field.FieldName);
  WriteLn(field.Value);
end.
"#,
    );
    assert_eq!(out, vec!["Price", "19.99"]);
}

#[test]
fn test_variant_clear_to_unassigned() {
    let out = run_pascal(
        r#"
program Test;
var v: Variant;
begin
  v := 'Active';
  WriteLn(VarIsEmpty(v));
  VarClear(v);
  WriteLn(VarIsEmpty(v));
end.
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_variant_vartostr_def() {
    let out = run_pascal(
        r#"
program Test;
var vNull: Variant;
begin
  vNull := Null;
  WriteLn(VarToStrDef(vNull, 'DefaultString'));
end.
"#,
    );
    assert_eq!(out, vec!["DefaultString"]);
}

#[test]
fn test_variant_varisnumeric_check() {
    let out = run_pascal(
        r#"
program Test;
var vNum, vStr: Variant;
begin
  vNum := 45.6;
  vStr := 'hello';
  WriteLn(VarIsNumeric(vNum));
  WriteLn(VarIsNumeric(vStr));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_variant_varisstr_check() {
    let out = run_pascal(
        r#"
program Test;
var vNum, vStr: Variant;
begin
  vNum := 10;
  vStr := 'text';
  WriteLn(VarIsStr(vNum));
  WriteLn(VarIsStr(vStr));
end.
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_variant_array_of_heterogeneous() {
    let out = run_pascal(
        r#"
program Test;
var vArr: Variant;
begin
  vArr := VarArrayOf([100, 'StringInArray', 3.14]);
  WriteLn(vArr[0]);
  WriteLn(vArr[1]);
  WriteLn(vArr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["100", "StringInArray", "3.14"]);
}

#[test]
fn test_variant_negation_operator() {
    let out = run_pascal(
        r#"
program Test;
var v, res: Variant;
begin
  v := 50;
  res := -v;
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["-50"]);
}
