use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 99: GUID Data Structures, String Conversions & IID Reflection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_guid_structure_sizeof() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(SizeOf(TGUID) = 16);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_string_round_trip() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g1, g2: TGUID; s: String;
begin
  s := '{12345678-1234-1234-1234-1234567890AB}';
  g1 := StringToGUID(s);
  g2 := StringToGUID(GUIDToString(g1));
  WriteLn(IsEqualGUID(g1, g2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_guid_tostring_format_length() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID; s: String;
begin
  g := StringToGUID('{87654321-4321-4321-4321-BA0987654321}');
  s := GUIDToString(g);
  WriteLn(Length(s) = 38);
  WriteLn(s[1] = '{');
  WriteLn(s[38] = '}');
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE", "TRUE"]);
}

#[test]
fn test_guid_equality_operators() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g1, g2, g3: TGUID;
begin
  g1 := StringToGUID('{11111111-2222-3333-4444-555555555555}');
  g2 := StringToGUID('{11111111-2222-3333-4444-555555555555}');
  g3 := StringToGUID('{99999999-8888-7777-6666-555555555555}');
  WriteLn(g1 = g2);
  WriteLn(g1 <> g3);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_guid_createguid_function() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID; res: Integer;
begin
  res := CreateGUID(g);
  WriteLn(res = 0);
  WriteLn(Length(GUIDToString(g)) = 38);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_guid_interface_iid_query() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type IMyCustomIntf = interface
  ['{AAAABBBB-CCCC-DDDD-EEEE-FFFF00001111}']
end;
var g: TGUID;
begin
  g := IMyCustomIntf;
  WriteLn(GUIDToString(g) = '{AAAABBBB-CCCC-DDDD-EEEE-FFFF00001111}');
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_null_check() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  FillChar(g, SizeOf(TGUID), 0);
  WriteLn(IsEqualGUID(g, GUID_NULL));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_guid_fields_direct_access() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  g.D1 := $12345678;
  g.D2 := $ABCD;
  g.D3 := $EF01;
  g.D4[0] := $11; g.D4[1] := $22;
  WriteLn(HexStr(g.D1, 8));
  WriteLn(HexStr(g.D2, 4));
end.
"#,
    );
    assert_eq!(out, vec!["12345678", "ABCD"]);
}

#[test]
fn test_guid_in_record_struct() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TComponentMeta = record
  ID: TGUID;
  Name: String;
end;
var meta: TComponentMeta;
begin
  meta.ID := StringToGUID('{00000000-0000-0000-0000-000000000001}');
  meta.Name := 'Component1';
  WriteLn(GUIDToString(meta.ID) + ':' + meta.Name);
end.
"#,
    );
    assert_eq!(
        out,
        vec!["{00000000-0000-0000-0000-000000000001}:Component1"]
    );
}

#[test]
fn test_guid_array_operations() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var guids: array[0..1] of TGUID;
begin
  guids[0] := StringToGUID('{10000000-0000-0000-0000-000000000000}');
  guids[1] := StringToGUID('{20000000-0000-0000-0000-000000000000}');
  WriteLn(IsEqualGUID(guids[0], guids[1]));
end.
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_guid_invalid_string_raises_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  try
    g := StringToGUID('InvalidGUIDStringFormat');
  except
    on E: EConvertError do WriteLn('InvalidGuidErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["InvalidGuidErrorCaught"]);
}

#[test]
fn test_guid_compare_equal() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g1, g2: TGUID;
begin
  g1 := StringToGUID('{ABCDEF01-1234-5678-90AB-CDEF01234567}');
  g2 := StringToGUID('{ABCDEF01-1234-5678-90AB-CDEF01234567}');
  WriteLn(g1 = g2);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_createguid_uniqueness() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g1, g2: TGUID;
begin
  CreateGUID(g1);
  CreateGUID(g2);
  WriteLn(not IsEqualGUID(g1, g2));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_lowercase_string_parsing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  g := StringToGUID('{abcdef01-1234-5678-90ab-cdef01234567}');
  WriteLn(GUIDToString(g) = '{ABCDEF01-1234-5678-90AB-CDEF01234567}');
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_without_braces_parsing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  g := StringToGUID('12345678-1234-1234-1234-1234567890AB');
  WriteLn(GUIDToString(g) = '{12345678-1234-1234-1234-1234567890AB}');
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_dynamic_array_search() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
function FindGUID(const list: array of TGUID; const target: TGUID): Boolean;
var i: Integer;
begin
  Result := False;
  for i := Low(list) to High(list) do
    if IsEqualGUID(list[i], target) then Exit(True);
end;

var list: array[0..1] of TGUID; target: TGUID;
begin
  list[0] := StringToGUID('{11111111-1111-1111-1111-111111111111}');
  list[1] := StringToGUID('{22222222-2222-2222-2222-222222222222}');
  target := StringToGUID('{22222222-2222-2222-2222-222222222222}');
  WriteLn(FindGUID(list, target));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_guid_hash_code_generation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
function HashGUID(const g: TGUID): Cardinal;
begin
  Result := g.D1 xor (LongWord(g.D2) shl 16 or g.D3);
end;
var g: TGUID;
begin
  g := StringToGUID('{12345678-0000-0000-0000-000000000000}');
  WriteLn(HashGUID(g) = $12345678);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_guid_d4_bytes_copy() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID; i: Integer; sum: Integer;
begin
  g := StringToGUID('{00000000-0000-0000-0102-030405060708}');
  sum := 0;
  for i := 0 to 7 do
    sum := sum + g.D4[i];
  WriteLn(sum); // 1+2+3+4+5+6+7+8 = 36
end.
"#,
    );
    assert_eq!(out, vec!["36"]);
}

#[test]
fn test_guid_supports_guid_variable() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var g: TGUID;
begin
  g := GUID_NULL;
  WriteLn(IsEqualGUID(g, GUID_NULL));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_guid_to_string_conversion_in_log() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure LogGuid(const g: TGUID);
begin
  WriteLn('LOG_GUID:' + GUIDToString(g));
end;
begin
  LogGuid(StringToGUID('{99999999-9999-9999-9999-999999999999}'));
end.
"#,
    );
    assert_eq!(out, vec!["LOG_GUID:{99999999-9999-9999-9999-999999999999}"]);
}
