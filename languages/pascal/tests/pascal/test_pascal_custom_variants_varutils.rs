use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 91: Custom Variants, VarUtils & Variant Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_variant_vartype_inspection() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 100;
  WriteLn(VarType(v) = varInteger);
  v := 'StringVal';
  WriteLn(VarType(v) = varString);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_variant_varisempty_varisnull() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := Unassigned;
  v2 := Null;
  WriteLn(VarIsEmpty(v1));
  WriteLn(VarIsNull(v2));
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_variant_varclear_resets_unassigned() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 'Data';
  VarClear(v);
  WriteLn(VarIsEmpty(v));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisstr_varisnumeric() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := 'Text'; v2 := 45.67;
  WriteLn(VarIsStr(v1));
  WriteLn(VarIsNumeric(v2));
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_variant_vartostr_vartostrdef() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 777;
  WriteLn(VarToStr(v));
  v := Null;
  WriteLn(VarToStrDef(v, 'DefaultNullText'));
end.
"#,
    );
    assert_eq!(out, vec!["777", "DefaultNullText"]);
}

#[test]
fn test_variant_varastype_explicit_coercion() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant; i: Integer;
begin
  v := '1234';
  v := VarAsType(v, varInteger);
  i := v;
  WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn test_variant_vartodatetime() {
    let out = run_pascal(
        r#"
program Test;
uses Variants, SysUtils;
var v: Variant; dt: TDateTime; y, m, d: Word;
begin
  v := '2026-10-15';
  dt := VarToDateTime(v);
  DecodeDate(dt, y, m, d);
  WriteLn(y.ToString + '-' + m.ToString + '-' + d.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["2026-10-15"]);
}

#[test]
fn test_variant_vartypetoasstring() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
begin
  WriteLn(VarTypeToAsString(varInteger));
  WriteLn(VarTypeToAsString(varString));
end.
"#,
    );
    assert_eq!(out, vec!["Integer", "String"]);
}

#[test]
fn test_variant_custom_variant_type_subclassing() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
type TCustomVarType = class(TCustomVariantType)
  public procedure Clear(var V: TVarData); override;
end;
procedure TCustomVarType.Clear(var V: TVarData);
begin
  V.VType := varEmpty;
end;

var cvt: TCustomVarType;
begin
  cvt := TCustomVarType.Create;
  WriteLn(cvt <> nil);
  cvt.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisbool() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := True;
  WriteLn(VarIsBool(v));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisarray() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := VarArrayCreate([0, 2], varInteger);
  WriteLn(VarIsArray(v));
  VarClear(v);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_arithmetic_operations() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2, res: Variant;
begin
  v1 := 10; v2 := 20;
  res := v1 + v2;
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_variant_comparison_operations() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := 50; v2 := 100;
  WriteLn(v1 < v2);
  WriteLn(v1 = v2);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_variant_string_concatenation() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2, res: Variant;
begin
  v1 := 'Hello '; v2 := 'Variant';
  res := v1 + v2;
  WriteLn(res);
end.
"#,
    );
    assert_eq!(out, vec!["Hello Variant"]);
}

#[test]
fn test_variant_null_propagation_in_math() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2, res: Variant;
begin
  v1 := 10; v2 := Null;
  res := v1 + v2;
  WriteLn(VarIsNull(res));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisbyref() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 100;
  WriteLn(VarIsByRef(v));
end.
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_variant_samevalue_check() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v1, v2: Variant;
begin
  v1 := 'Same'; v2 := 'Same';
  WriteLn(VarSameValue(v1, v2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisordinal() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 42;
  WriteLn(VarIsOrdinal(v));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varisfloat() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;
var v: Variant;
begin
  v := 3.14159;
  WriteLn(VarIsFloat(v));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_variant_varcast_exception() {
    let out = run_pascal(
        r#"
program Test;
uses Variants, SysUtils;
var v: Variant;
begin
  v := 'NotANumberString';
  try
    v := VarAsType(v, varInteger);
  except
    on E: EVariantError do WriteLn('VariantCastFailed');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["VariantCastFailed"]);
}
