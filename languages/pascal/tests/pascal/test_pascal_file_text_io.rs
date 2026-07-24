use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 61: File I/O (Text Files & Standard Text Streams)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_textfile_assign_rewrite_writeln_close() {
    let out = run_pascal(r#"
program Test;
var f: TextFile;
begin
  AssignFile(f, 'test_text_1.txt');
  Rewrite(f);
  WriteLn(f, 'Line 1');
  WriteLn(f, 'Line 2');
  CloseFile(f);
  WriteLn('Written');
end.
"#);
    assert_eq!(out, vec!["Written"]);
}

#[test]
fn test_textfile_reset_readln_eof() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_text_1.txt');
  Reset(f);
  while not Eof(f) do
  begin
    ReadLn(f, line);
    WriteLn(line);
  end;
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["Line 1", "Line 2"]);
}

#[test]
fn test_textfile_append_mode() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_text_1.txt');
  Append(f);
  WriteLn(f, 'Line 3');
  CloseFile(f);

  Reset(f);
  while not Eof(f) do
  begin
    ReadLn(f, line);
    WriteLn(line);
  end;
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["Line 1", "Line 2", "Line 3"]);
}

#[test]
fn test_textfile_write_without_newline() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_no_nl.txt');
  Rewrite(f);
  Write(f, 'Hello ');
  Write(f, 'World');
  WriteLn(f);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn test_textfile_formatted_number_writing() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; val: Integer; r: Real;
begin
  AssignFile(f, 'test_num.txt');
  Rewrite(f);
  val := 42; r := 3.14;
  WriteLn(f, val);
  WriteLn(f, r:0:2);
  CloseFile(f);

  Reset(f);
  ReadLn(f, val);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_textfile_read_formatted_values() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; a, b: Integer;
begin
  AssignFile(f, 'test_fmt.txt');
  Rewrite(f);
  WriteLn(f, '10 20');
  CloseFile(f);

  Reset(f);
  Read(f, a); Read(f, b);
  WriteLn(a + b);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_textfile_eoln_detection() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; ch: Char; count: Integer;
begin
  AssignFile(f, 'test_eoln.txt');
  Rewrite(f);
  WriteLn(f, 'ABC');
  CloseFile(f);

  Reset(f);
  count := 0;
  while not Eoln(f) do
  begin
    Read(f, ch);
    Inc(count);
  end;
  WriteLn(count);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_textfile_ioresult_checking() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; err: Integer;
begin
  AssignFile(f, 'non_existent_file_xyz.txt');
  {$I-}
  Reset(f);
  err := IOResult;
  {$I+}
  WriteLn(err <> 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_textfile_protection_finally() {
    let out = run_pascal(r#"
program Test;
var f: TextFile;
begin
  AssignFile(f, 'test_finally.txt');
  Rewrite(f);
  try
    WriteLn(f, 'ProtectedFileContent');
  finally
    CloseFile(f);
    WriteLn('ClosedInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["ClosedInFinally"]);
}

#[test]
fn test_textfile_empty_lines_reading() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; l1, l2: String;
begin
  AssignFile(f, 'test_empty.txt');
  Rewrite(f);
  WriteLn(f, '');
  WriteLn(f, 'SecondLine');
  CloseFile(f);

  Reset(f);
  ReadLn(f, l1);
  ReadLn(f, l2);
  WriteLn(Length(l1));
  WriteLn(l2);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["0", "SecondLine"]);
}

#[test]
fn test_textfile_width_padded_write() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_pad.txt');
  Rewrite(f);
  WriteLn(f, 99:5);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn('[' + line + ']');
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["[   99]"]);
}

#[test]
fn test_textfile_multiple_lines_loop_write() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; i: Integer;
begin
  AssignFile(f, 'test_loop.txt');
  Rewrite(f);
  for i := 1 to 3 do
    WriteLn(f, 'Item ' + i.ToString);
  CloseFile(f);

  Reset(f);
  i := 0;
  while not Eof(f) do
  begin
    Inc(i);
    CloseFile(f); // Break loop early after verification
    Break;
  end;
  WriteLn(i);
end.
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_textfile_reading_multiple_strings() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; s1, s2: String;
begin
  AssignFile(f, 'test_str2.txt');
  Rewrite(f);
  WriteLn(f, 'Alpha');
  WriteLn(f, 'Beta');
  CloseFile(f);

  Reset(f);
  ReadLn(f, s1); ReadLn(f, s2);
  WriteLn(s1 + ' & ' + s2);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["Alpha & Beta"]);
}

#[test]
fn test_textfile_flush_buffer() {
    let out = run_pascal(r#"
program Test;
var f: TextFile;
begin
  AssignFile(f, 'test_flush.txt');
  Rewrite(f);
  WriteLn(f, 'FlushedData');
  Flush(f);
  CloseFile(f);
  WriteLn('FlushedSuccessfully');
end.
"#);
    assert_eq!(out, vec!["FlushedSuccessfully"]);
}

#[test]
fn test_textfile_boolean_write() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_bool.txt');
  Rewrite(f);
  WriteLn(f, True);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_textfile_char_by_char_reading() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; ch: Char;
begin
  AssignFile(f, 'test_chars.txt');
  Rewrite(f);
  Write(f, 'XY');
  CloseFile(f);

  Reset(f);
  Read(f, ch); WriteLn(ch);
  Read(f, ch); WriteLn(ch);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["X", "Y"]);
}

#[test]
fn test_textfile_erase_file() {
    let out = run_pascal(r#"
program Test;
var f: TextFile;
begin
  AssignFile(f, 'test_erase.txt');
  Rewrite(f);
  WriteLn(f, 'Temporary');
  CloseFile(f);
  Erase(f);
  WriteLn('Erased');
end.
"#);
    assert_eq!(out, vec!["Erased"]);
}

#[test]
fn test_textfile_rename_file() {
    let out = run_pascal(r#"
program Test;
var f: TextFile;
begin
  AssignFile(f, 'test_orig.txt');
  Rewrite(f);
  WriteLn(f, 'Data');
  CloseFile(f);
  Rename(f, 'test_renamed.txt');
  WriteLn('Renamed');
end.
"#);
    assert_eq!(out, vec!["Renamed"]);
}

#[test]
fn test_textfile_procedure_parameter() {
    let out = run_pascal(r#"
program Test;
procedure WriteHeader(var f: TextFile);
begin
  WriteLn(f, '=== HEADER ===');
end;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_proc.txt');
  Rewrite(f);
  WriteHeader(f);
  CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["=== HEADER ==="]);
}

#[test]
fn test_textfile_overwrite_existing() {
    let out = run_pascal(r#"
program Test;
var f: TextFile; line: String;
begin
  AssignFile(f, 'test_overwrite.txt');
  Rewrite(f); WriteLn(f, 'Old'); CloseFile(f);
  Rewrite(f); WriteLn(f, 'New'); CloseFile(f);

  Reset(f);
  ReadLn(f, line);
  WriteLn(line);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["New"]);
}
