use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 62: Typed Binary Files (file of TRecord & Seek/FilePos)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_typed_file_integer_write_read() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_int.bin');
  Rewrite(f);
  val := 100; Write(f, val);
  val := 200; Write(f, val);
  CloseFile(f);

  Reset(f);
  Read(f, val); WriteLn(val);
  Read(f, val); WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_typed_file_record_write_read() {
    let out = run_pascal(r#"
program Test;
type TCustomer = packed record
  ID: Integer;
  Score: Word;
end;
var f: file of TCustomer; c1, c2: TCustomer;
begin
  AssignFile(f, 'test_cust.bin');
  Rewrite(f);
  c1.ID := 1; c1.Score := 95; Write(f, c1);
  c1.ID := 2; c1.Score := 88; Write(f, c1);
  CloseFile(f);

  Reset(f);
  Read(f, c2); WriteLn(c2.ID.ToString + ':' + c2.Score.ToString);
  Read(f, c2); WriteLn(c2.ID.ToString + ':' + c2.Score.ToString);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["1:95", "2:88"]);
}

#[test]
fn test_typed_file_filesize_and_filepos() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_pos.bin');
  Rewrite(f);
  val := 10; Write(f, val);
  val := 20; Write(f, val);
  val := 30; Write(f, val);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f));
  WriteLn(FilePos(f));
  Read(f, val);
  WriteLn(FilePos(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["3", "0", "1"]);
}

#[test]
fn test_typed_file_seek_random_access() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_seek.bin');
  Rewrite(f);
  val := 10; Write(f, val);
  val := 20; Write(f, val);
  val := 30; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_typed_file_append_using_seek_filesize() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_app.bin');
  Rewrite(f);
  val := 1; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, FileSize(f));
  val := 2; Write(f, val);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_typed_file_in_place_update() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_update.bin');
  Rewrite(f);
  val := 100; Write(f, val);
  val := 200; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  val := val + 50;
  Seek(f, 1);
  Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["250"]);
}

#[test]
fn test_typed_file_truncate_at_position() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_trunc.bin');
  Rewrite(f);
  val := 1; Write(f, val);
  val := 2; Write(f, val);
  val := 3; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, 1);
  Truncate(f);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_typed_file_byte_stream() {
    let out = run_pascal(r#"
program Test;
var f: file of Byte; b: Byte;
begin
  AssignFile(f, 'test_byte.bin');
  Rewrite(f);
  b := $AB; Write(f, b);
  b := $CD; Write(f, b);
  CloseFile(f);

  Reset(f);
  Read(f, b); WriteLn(HexStr(b, 2));
  Read(f, b); WriteLn(HexStr(b, 2));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["AB", "CD"]);
}

#[test]
fn test_typed_file_word_stream() {
    let out = run_pascal(r#"
program Test;
var f: file of Word; w: Word;
begin
  AssignFile(f, 'test_word.bin');
  Rewrite(f);
  w := 65000; Write(f, w);
  CloseFile(f);

  Reset(f);
  Read(f, w); WriteLn(w);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["65000"]);
}

#[test]
fn test_typed_file_double_stream() {
    let out = run_pascal(r#"
program Test;
var f: file of Real; r: Real;
begin
  AssignFile(f, 'test_real.bin');
  Rewrite(f);
  r := 12.34; Write(f, r);
  CloseFile(f);

  Reset(f);
  Read(f, r); WriteLn(r);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["12.34"]);
}

#[test]
fn test_typed_file_eof_iteration() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val, sum: Integer;
begin
  AssignFile(f, 'test_sum.bin');
  Rewrite(f);
  val := 5; Write(f, val);
  val := 15; Write(f, val);
  val := 25; Write(f, val);
  CloseFile(f);

  Reset(f);
  sum := 0;
  while not Eof(f) do
  begin
    Read(f, val);
    sum := sum + val;
  end;
  WriteLn(sum);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["45"]);
}

#[test]
fn test_typed_file_enum_type() {
    let out = run_pascal(r#"
program Test;
type TStatus = (stInit, stRunning, stDone);
var f: file of TStatus; s: TStatus;
begin
  AssignFile(f, 'test_enum.bin');
  Rewrite(f);
  s := stRunning; Write(f, s);
  CloseFile(f);

  Reset(f);
  Read(f, s); WriteLn(Ord(s));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_typed_file_protection_finally() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_prot.bin');
  Rewrite(f);
  try
    val := 777; Write(f, val);
  finally
    CloseFile(f);
    WriteLn('ClosedTypedFileInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["ClosedTypedFileInFinally"]);
}

#[test]
fn test_typed_file_seek_last_element() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_last.bin');
  Rewrite(f);
  val := 10; Write(f, val);
  val := 99; Write(f, val);
  CloseFile(f);

  Reset(f);
  Seek(f, FileSize(f) - 1);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_typed_file_procedure_parameter() {
    let out = run_pascal(r#"
program Test;
procedure WriteInt(var f: file of Integer; v: Integer);
begin
  Write(f, v);
end;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_proc_bin.bin');
  Rewrite(f);
  WriteInt(f, 888);
  CloseFile(f);

  Reset(f);
  Read(f, val);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["888"]);
}

#[test]
fn test_typed_file_boolean_type() {
    let out = run_pascal(r#"
program Test;
var f: file of Boolean; b: Boolean;
begin
  AssignFile(f, 'test_bool.bin');
  Rewrite(f);
  b := True; Write(f, b);
  b := False; Write(f, b);
  CloseFile(f);

  Reset(f);
  Read(f, b); WriteLn(b);
  Read(f, b); WriteLn(b);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_typed_file_empty_check() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer;
begin
  AssignFile(f, 'test_empty.bin');
  Rewrite(f);
  CloseFile(f);

  Reset(f);
  WriteLn(FileSize(f) = 0);
  WriteLn(Eof(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_typed_file_erase() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_erase.bin');
  Rewrite(f);
  val := 42; Write(f, val);
  CloseFile(f);
  Erase(f);
  WriteLn('ErasedTypedFile');
end.
"#);
    assert_eq!(out, vec!["ErasedTypedFile"]);
}

#[test]
fn test_typed_file_rename() {
    let out = run_pascal(r#"
program Test;
var f: file of Integer; val: Integer;
begin
  AssignFile(f, 'test_orig.bin');
  Rewrite(f);
  val := 55; Write(f, val);
  CloseFile(f);
  Rename(f, 'test_renamed.bin');
  WriteLn('RenamedTypedFile');
end.
"#);
    assert_eq!(out, vec!["RenamedTypedFile"]);
}

#[test]
fn test_typed_file_array_struct_payload() {
    let out = run_pascal(r#"
program Test;
type TArrayPayload = packed record
  Data: array[0..2] of Integer;
end;
var f: file of TArrayPayload; p1, p2: TArrayPayload;
begin
  AssignFile(f, 'test_arr_payload.bin');
  Rewrite(f);
  p1.Data[0] := 1; p1.Data[1] := 2; p1.Data[2] := 3;
  Write(f, p1);
  CloseFile(f);

  Reset(f);
  Read(f, p2);
  WriteLn(p2.Data[0] + p2.Data[1] + p2.Data[2]);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["6"]);
}
