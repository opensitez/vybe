use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 39: TStringBuilder Performance Buffer & Utilities
// ═══════════════════════════════════════════════════════════

#[test]
fn test_stringbuilder_basic_append() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('Hello ');
  sb.Append('World');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn test_stringbuilder_append_multiple_types() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('Int: ');
  sb.Append(42);
  sb.Append(' Bool: ');
  sb.Append(True);
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Int: 42 Bool: True"]);
}

#[test]
fn test_stringbuilder_method_chaining() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('A').Append('B').Append('C');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn test_stringbuilder_appendline() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.AppendLine('Line1');
  sb.Append('Line2');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Line1", "Line2"]);
}

#[test]
fn test_stringbuilder_insert_at_index() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('AC');
  sb.Insert(1, 'B');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn test_stringbuilder_remove_range() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('Hello World');
  sb.Remove(5, 6);
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn test_stringbuilder_replace_occurrences() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('foo bar foo');
  sb.Replace('foo', 'qux');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["qux bar qux"]);
}

#[test]
fn test_stringbuilder_clear_reset() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('Populated');
  sb.Clear;
  WriteLn(sb.Length);
  sb.Append('FreshData');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0", "FreshData"]);
}

#[test]
fn test_stringbuilder_appendformat() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.AppendFormat('Item: %s, ID: %d', ['Book', 101]);
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Item: Book, ID: 101"]);
}

#[test]
fn test_stringbuilder_chars_0based_indexing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('Pascal');
  WriteLn(sb.Chars[0]);
  sb.Chars[0] := 'J';
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["P", "Jascal"]);
}

#[test]
fn test_stringbuilder_capacity_preallocation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create(1024);
  WriteLn(sb.Capacity >= 1024);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_stringbuilder_loop_aggregation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder; i: Integer;
begin
  sb := TStringBuilder.Create;
  for i := 1 to 5 do
  begin
    if i > 1 then sb.Append(',');
    sb.Append(i);
  end;
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1,2,3,4,5"]);
}

#[test]
fn test_stringbuilder_json_formatting() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('{');
  sb.Append('"id":').Append(10).Append(',');
  sb.Append('"active":').Append('true');
  sb.Append('}');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["{\"id\":10,\"active\":true}"]);
}

#[test]
fn test_stringbuilder_xml_tag_building() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('<title>').Append('Pascal Documentation').Append('</title>');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["<title>Pascal Documentation</title>"]);
}

#[test]
fn test_stringbuilder_parameter_passing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure BuildHeader(sb: TStringBuilder);
begin
  sb.Append('=== HEADER ===');
end;
var b: TStringBuilder;
begin
  b := TStringBuilder.Create;
  BuildHeader(b);
  WriteLn(b.ToString);
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["=== HEADER ==="]);
}

#[test]
fn test_stringbuilder_append_float() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append(3.14);
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn test_stringbuilder_length_truncation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('LongStringBuilderText');
  sb.Length := 4;
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Long"]);
}

#[test]
fn test_stringbuilder_append_char() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('X');
  sb.Append('Y');
  sb.Append('Z');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["XYZ"]);
}

#[test]
fn test_stringbuilder_replace_single_char() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('A-B-C');
  sb.Replace('-', ':');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["A:B:C"]);
}

#[test]
fn test_stringbuilder_initial_string_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('InitialContent');
  sb.Append('_Appended');
  WriteLn(sb.ToString);
  sb.Free;
end.
"#,
    );
    assert_eq!(out, vec!["InitialContent_Appended"]);
}
