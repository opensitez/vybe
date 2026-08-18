use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 66: String Streams (TStringStream & DataString Access)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tstringstream_initial_datastring() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('InitialStringText');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["InitialStringText"]);
}

#[test]
fn test_tstringstream_writestring_append() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  ss.WriteString('Hello ');
  ss.WriteString('World');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn test_tstringstream_readstring_slice() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('Hello World');
  WriteLn(ss.ReadString(5));
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn test_tstringstream_position_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('Pascal');
  WriteLn(ss.Position);
  ss.Position := 3;
  WriteLn(ss.ReadString(3));
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0", "cal"]);
}

#[test]
fn test_tstringstream_size_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('Data');
  WriteLn(ss.Size);
  ss.Size := 2;
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["4", "Da"]);
}

#[test]
fn test_tstringstream_clear_content() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('Populated');
  ss.Size := 0;
  WriteLn(Length(ss.DataString));
  ss.WriteString('Fresh');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0", "Fresh"]);
}

#[test]
fn test_tstringstream_copy_to_memorystream() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream; ms: TMemoryStream;
begin
  ss := TStringStream.Create('StreamCopyText');
  ms := TMemoryStream.Create;
  ms.CopyFrom(ss, ss.Size);
  WriteLn(ms.Size);
  ms.Free; ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn test_tstringstream_multiline_text() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  ss.WriteString('Line1' + #13#10);
  ss.WriteString('Line2');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Line1", "Line2"]);
}

#[test]
fn test_tstringstream_polymorphic_parameter() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
procedure AppendHeader(stream: TStream);
var text: String;
begin
  text := '[HEADER]';
  stream.WriteBuffer(text[1], Length(text));
end;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  AppendHeader(ss);
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["[HEADER]"]);
}

#[test]
fn test_tstringstream_protection_finally() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('ProtectedStreamText');
  try
    WriteLn(ss.DataString);
  finally
    ss.Free;
    WriteLn('StringStreamFreed');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ProtectedStreamText", "StringStreamFreed"]);
}

#[test]
fn test_tstringstream_append_formatted_numbers() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  ss.WriteString(Format('Val: %d, Rate: %.2f', [100, 5.5]));
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Val: 100, Rate: 5.50"]);
}

#[test]
fn test_tstringstream_json_building() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  ss.WriteString('{"status":"ok"}');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["{\"status\":\"ok\"}"]);
}

#[test]
fn test_tstringstream_xml_building() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create;
  ss.WriteString('<root><item>1</item></root>');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["<root><item>1</item></root>"]);
}

#[test]
fn test_tstringstream_seek_sofromend() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('Text');
  ss.Seek(0, soFromEnd);
  WriteLn(ss.Position);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_tstringstream_empty_initial_text() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('');
  WriteLn(ss.Size);
  WriteLn(ss.DataString = '');
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0", "TRUE"]);
}

#[test]
fn test_tstringstream_overwrite_middle_character() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('A_C');
  ss.Position := 1;
  ss.WriteString('B');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn test_tstringstream_loop_string_building() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream; i: Integer;
begin
  ss := TStringStream.Create;
  for i := 1 to 3 do
  begin
    if i > 1 then ss.WriteString('-');
    ss.WriteString(i.ToString);
  end;
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1-2-3"]);
}

#[test]
fn test_tstringstream_readbuffer_into_char_array() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream; chars: array[0..2] of Char;
begin
  ss := TStringStream.Create('XYZ');
  ss.ReadBuffer(chars[0], 3);
  WriteLn(chars[0] + chars[1] + chars[2]);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["XYZ"]);
}

#[test]
fn test_tstringstream_utf8_encoding_creation() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ss: TStringStream;
begin
  ss := TStringStream.Create('UTF8StringText', TEncoding.UTF8);
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["UTF8StringText"]);
}

#[test]
fn test_tstringstream_datastring_reassignment() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ss: TStringStream;
begin
  ss := TStringStream.Create('First');
  ss.Size := 0;
  ss.WriteString('Second');
  WriteLn(ss.DataString);
  ss.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Second"]);
}
