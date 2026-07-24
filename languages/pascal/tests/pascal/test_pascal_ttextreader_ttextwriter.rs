use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 68: Text Stream Readers & Writers (TStreamReader/Writer)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tstreamwriter_and_tstreamreader_basic_lines() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('Line 1');
  w.WriteLine('Line 2');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["Line 1", "Line 2"]);
}

#[test]
fn test_tstreamreader_readtoend() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('FirstLine');
  w.WriteLine('SecondLine');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadToEnd);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["FirstLine", "SecondLine"]);
}

#[test]
fn test_tstreamreader_endofstream_check() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('SingleLine');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.EndOfStream);
  r.ReadLine;
  WriteLn(r.EndOfStream);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_tstreamwriter_autoflush_enabled() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.AutoFlush := True;
  w.WriteLine('AutoFlushedText');
  WriteLn(ms.Size > 0);
  w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tstreamreader_peek_next_char() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('ABC');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.Peek = Ord('A'));
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["True", "ABC"]);
}

#[test]
fn test_tstreamwriter_utf8_encoding() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms, TEncoding.UTF8);
  w.WriteLine('UTF8EncodedContent');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms, TEncoding.UTF8);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["UTF8EncodedContent"]);
}

#[test]
fn test_tstreamwriter_formatted_numbers() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine(Format('Count: %d, Value: %.2f', [42, 99.95]));
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["Count: 42, Value: 99.95"]);
}

#[test]
fn test_tstreamreader_loop_line_count() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader; count: Integer;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('1'); w.WriteLine('2'); w.WriteLine('3');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  count := 0;
  while not r.EndOfStream do
  begin
    r.ReadLine;
    Inc(count);
  end;
  WriteLn(count);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_tstreamwriter_write_without_newline() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.Write('Part1_');
  w.Write('Part2');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["Part1_Part2"]);
}

#[test]
fn test_tstreamwriter_with_tfilestream() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var fs: TFileStream; w: TStreamWriter; r: TStreamReader;
begin
  fs := TFileStream.Create('test_text_rw.txt', fmCreate);
  w := TStreamWriter.Create(fs);
  w.WriteLine('FileStreamText');
  w.Free;

  fs := TFileStream.Create('test_text_rw.txt', fmOpenRead);
  r := TStreamReader.Create(fs);
  WriteLn(r.ReadLine);
  r.Free;
end.
"#);
    assert_eq!(out, vec!["FileStreamText"]);
}

#[test]
fn test_tstreamwriter_protection_finally() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  try
    w.WriteLine('ProtectedTextWrite');
  finally
    w.Free;
    ms.Free;
    WriteLn('StreamWriterFreedInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["StreamWriterFreedInFinally"]);
}

#[test]
fn test_tstreamreader_empty_stream() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  r := TStreamReader.Create(ms);
  WriteLn(r.EndOfStream);
  WriteLn(Length(r.ReadLine));
  r.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["True", "0"]);
}

#[test]
fn test_tstreamwriter_csv_formatting() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('1,Alice,100');
  w.WriteLine('2,Bob,200');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["1,Alice,100", "2,Bob,200"]);
}

#[test]
fn test_tstreamwriter_json_formatting() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('{');
  w.WriteLine('  "name": "Pascal"');
  w.WriteLine('}');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["{", "  \"name\": \"Pascal\""]);
}

#[test]
fn test_tstreamwriter_xml_formatting() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine('<xml>');
  w.WriteLine('  <val>10</val>');
  w.WriteLine('</xml>');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["<xml>", "  <val>10</val>"]);
}

#[test]
fn test_tstreamreader_read_char_by_char() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.Write('XY');
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(Chr(r.Read));
  WriteLn(Chr(r.Read));
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["X", "Y"]);
}

#[test]
fn test_tstreamwriter_write_boolean() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine(True);
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tstreamwriter_write_integer() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine(12345);
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn test_tstreamwriter_write_float() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TStreamWriter; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  w := TStreamWriter.Create(ms);
  w.WriteLine(3.14);
  w.Flush;

  ms.Position := 0;
  r := TStreamReader.Create(ms);
  WriteLn(r.ReadLine);
  r.Free; w.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn test_tstreamreader_base_stream_access() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; r: TStreamReader;
begin
  ms := TMemoryStream.Create;
  r := TStreamReader.Create(ms);
  WriteLn(r.BaseStream <> nil);
  r.Free; ms.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}
