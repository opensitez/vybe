use super::helpers::run_pascal;

#[test]
fn test_format_width_integer() {
    let src = r#"
program T;
begin
  WriteLn(Format('%5d', [42]));
  WriteLn(Format('%-5d', [42]));
  WriteLn(Format('%05d', [42]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["   42", "42   ", "00042"]);
}

#[test]
fn test_format_float_precision() {
    let src = r#"
program T;
begin
  WriteLn(Format('%.2f', [3.14159]));
  WriteLn(Format('%.4f', [3.14159]));
  WriteLn(Format('%8.2f', [3.14159]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3.14", "3.1416", "    3.14"]);
}

#[test]
fn test_format_string_width() {
    let src = r#"
program T;
begin
  WriteLn(Format('%-10s|', ['left']));
  WriteLn(Format('%10s|', ['right']));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["left      |", "     right|"]);
}

#[test]
fn test_format_hex_specifier() {
    let src = r#"
program T;
begin
  WriteLn(Format('%x', [255]));
  WriteLn(Format('%X', [255]));
  WriteLn(Format('%08x', [255]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["ff", "FF", "000000ff"]);
}

#[test]
fn test_format_scientific() {
    let src = r#"
program T;
begin
  WriteLn(Format('%e', [1234.5]));
  WriteLn(Format('%.2e', [1234.5]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1.2345e+03", "1.23e+03"]);
}

#[test]
fn test_format_multiple_mixed() {
    let src = r#"
program T;
begin
  WriteLn(Format('%s=%d (%.1f%%)', ['score', 85, 85.0]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["score=85 (85.0%)"]);
}

#[test]
fn test_format_table_row() {
    let src = r#"
program T;
procedure TableRow(name: string; val: Integer);
begin
  WriteLn(Format('%-12s %6d', [name, val]));
end;
begin
  TableRow('Alice', 100);
  TableRow('Bob', 85);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Alice          100", "Bob             85"]);
}

#[test]
fn test_str_procedure() {
    let src = r#"
program T;
var
  s: string;
  n: Integer;
  f: Double;
begin
  n := 42;
  Str(n, s);
  WriteLn(s);
  f := 3.14;
  Str(f:6:2, s);
  WriteLn(s);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42", "  3.14"]);
}

#[test]
fn test_val_procedure_integer() {
    let src = r#"
program T;
var
  n: Integer;
  code: Integer;
begin
  Val('123', n, code);
  WriteLn(n);
  WriteLn(code);
  Val('bad', n, code);
  WriteLn(code > 0);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["123", "0", "true"]);
}

#[test]
fn test_format_signed_integers() {
    let src = r#"
program T;
begin
  WriteLn(Format('%+d', [42]));
  WriteLn(Format('%+d', [-42]));
  WriteLn(Format('%d', [-100]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["+42", "-42", "-100"]);
}

#[test]
fn test_writeln_width_specifier() {
    let src = r#"
program T;
var
  n: Integer;
  f: Double;
begin
  n := 42;
  WriteLn(n:8);
  f := 3.14;
  WriteLn(f:8:2);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["      42", "    3.14"]);
}

#[test]
fn test_write_specifier_aligned() {
    let src = r#"
program T;
begin
  Write(1:4);
  Write(2:4);
  Write(3:4);
  WriteLn('');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["   1   2   3"]);
}

#[test]
fn test_format_build_path() {
    let src = r#"
program T;
var
  dir, file, ext: string;
begin
  dir := '/home/user';
  file := 'document';
  ext := 'txt';
  WriteLn(Format('%s/%s.%s', [dir, file, ext]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["/home/user/document.txt"]);
}

#[test]
fn test_format_json_like() {
    let src = r#"
program T;
begin
  WriteLn(Format('{"name":"%s","age":%d}', ['Alice', 30]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["[{\"name\":\"Alice\",\"age\":30}]"]);
}

#[test]
fn test_format_percentage() {
    let src = r#"
program T;
var
  part, total: Integer;
  pct: Double;
begin
  part := 3;
  total := 4;
  pct := part / total * 100;
  WriteLn(Format('%.1f%%', [pct]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["75.0%"]);
}

#[test]
fn test_format_integer_zero_padded() {
    let src = r#"
program T;
var
  i: Integer;
begin
  for i := 1 to 5 do
    WriteLn(Format('item_%03d', [i]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(
        out,
        vec!["item_001", "item_002", "item_003", "item_004", "item_005"]
    );
}

#[test]
fn test_format_repeat_in_loop() {
    let src = r#"
program T;
var
  names: array[0..2] of string;
  vals: array[0..2] of Integer;
  i: Integer;
begin
  names[0] := 'x'; vals[0] := 10;
  names[1] := 'y'; vals[1] := 20;
  names[2] := 'z'; vals[2] := 30;
  for i := 0 to 2 do
    WriteLn(Format('%s: %d', [names[i], vals[i]]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["x: 10", "y: 20", "z: 30"]);
}

#[test]
fn test_format_negative_width() {
    let src = r#"
program T;
begin
  WriteLn(Format('|%-8s|', ['abc']));
  WriteLn(Format('|%-8d|', [42]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["|abc     |", "|42      |"]);
}

#[test]
fn test_format_combined_table() {
    let src = r#"
program T;
type
  TRow = record
    Name: string;
    Score: Integer;
    Pct: Double;
  end;
var
  rows: array[0..1] of TRow;
  i: Integer;
begin
  rows[0].Name := 'Alice'; rows[0].Score := 95; rows[0].Pct := 95.0;
  rows[1].Name := 'Bob';   rows[1].Score := 72; rows[1].Pct := 72.0;
  for i := 0 to 1 do
    WriteLn(Format('%-8s %3d  %.1f%%', [rows[i].Name, rows[i].Score, rows[i].Pct]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Alice     95  95.0%", "Bob       72  72.0%"]);
}

#[test]
fn test_format_g_specifier() {
    let src = r#"
program T;
begin
  WriteLn(Format('%g', [1000000.0]));
  WriteLn(Format('%g', [0.0001]));
  WriteLn(Format('%g', [3.14]));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1e+06", "0.0001", "3.14"]);
}
