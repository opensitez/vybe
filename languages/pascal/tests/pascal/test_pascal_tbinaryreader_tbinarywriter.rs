use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 67: Binary Stream Readers & Writers (TBinaryReader/Writer)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tbinarywriter_and_tbinaryreader_primitive_int() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(42);
  w.Write(100);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadInt32);
  WriteLn(r.ReadInt32);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42", "100"]);
}

#[test]
fn test_tbinarywriter_and_tbinaryreader_string() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write('BinaryStringPayload');
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadString);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["BinaryStringPayload"]);
}

#[test]
fn test_tbinarywriter_and_tbinaryreader_boolean() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(True);
  w.Write(False);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadBoolean);
  WriteLn(r.ReadBoolean);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "FALSE"]);
}

#[test]
fn test_tbinarywriter_and_tbinaryreader_double() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(3.14159);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadDouble);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn test_tbinarywriter_heterogeneous_pipeline() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(101);
  w.Write('ItemName');
  w.Write(True);
  w.Write(19.99);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadInt32);
  WriteLn(r.ReadString);
  WriteLn(r.ReadBoolean);
  WriteLn(r.ReadDouble);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["101", "ItemName", "TRUE", "19.99"]);
}

#[test]
fn test_tbinaryreader_readbytes() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
    bytes: TBytes;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(Byte(10)); w.Write(Byte(20)); w.Write(Byte(30));
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  bytes := r.ReadBytes(3);
  WriteLn(Length(bytes));
  WriteLn(bytes[0]);
  WriteLn(bytes[2]);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "10", "30"]);
}

#[test]
fn test_tbinarywriter_with_tfilestream() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var fs: TFileStream; w: TBinaryWriter; r: TBinaryReader;
begin
  fs := TFileStream.Create('test_bin_rw.dat', fmCreate);
  w := TBinaryWriter.Create(fs);
  w.Write('FileBinaryData');
  w.Free;

  fs := TFileStream.Create('test_bin_rw.dat', fmOpenRead);
  r := TBinaryReader.Create(fs);
  WriteLn(r.ReadString);
  r.Free;
end.
"#,
    );
    assert_eq!(out, vec!["FileBinaryData"]);
}

#[test]
fn test_tbinaryreader_basestream_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  r := TBinaryReader.Create(ms);
  WriteLn(r.BaseStream <> nil);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tbinarywriter_write_byte_word_cardinal() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(Byte(255));
  w.Write(Word(65535));
  w.Write(Cardinal(4294967295));
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadByte);
  WriteLn(r.ReadUInt16);
  WriteLn(r.ReadUInt32);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["255", "65535", "4294967295"]);
}

#[test]
fn test_tbinarywriter_write_char() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write('Z');
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadChar);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn test_tbinarywriter_protection_finally() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  try
    w.Write('ProtectedBinaryWrite');
  finally
    w.Free;
    ms.Free;
    WriteLn('BinaryWriterFreedInFinally');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["BinaryWriterFreedInFinally"]);
}

#[test]
fn test_tbinarywriter_loop_writing() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader; i: Integer;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  for i := 1 to 3 do w.Write(i * 10);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  for i := 1 to 3 do WriteLn(r.ReadInt32);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_tbinaryreader_readint64() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(Int64(9876543210));
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadInt64);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["9876543210"]);
}

#[test]
fn test_tbinaryreader_readsingle_float() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(Single(12.5));
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadSingle);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["12.5"]);
}

#[test]
fn test_tbinarywriter_empty_string_write() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write('');
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(Length(r.ReadString));
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tbinarywriter_multiple_strings() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write('First');
  w.Write('Second');
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadString + ' & ' + r.ReadString);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["First & Second"]);
}

#[test]
fn test_tbinaryreader_peekchar() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write('A');
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.PeekChar = Ord('A'));
  WriteLn(r.ReadChar);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "A"]);
}

#[test]
fn test_tbinarywriter_record_serialization() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
type THeader = packed record Magic: Word; Version: Byte; end;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader; h: THeader;
begin
  h.Magic := $4D5A; h.Version := 1;
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(h.Magic);
  w.Write(h.Version);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadUInt16);
  WriteLn(r.ReadByte);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["19802", "1"]);
}

#[test]
fn test_tbinarywriter_array_of_bytes_write() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
    writeBytes, readBytes: TBytes;
begin
  SetLength(writeBytes, 2);
  writeBytes[0] := 55; writeBytes[1] := 66;
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(writeBytes);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  readBytes := r.ReadBytes(2);
  WriteLn(readBytes[0]);
  WriteLn(readBytes[1]);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["55", "66"]);
}

#[test]
fn test_tbinaryreader_eof_check_base_stream() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;
var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(123);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  r.ReadInt32;
  WriteLn(r.BaseStream.Position = r.BaseStream.Size);
  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}
