use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 64: OOP Stream Abstractions (TStream & TFileStream)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tfilestream_create_and_writebuffer() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  fs := TFileStream.Create('test_stream_1.dat', fmCreate);
  try
    val := 500;
    fs.WriteBuffer(val, SizeOf(Integer));
  finally
    fs.Free;
  end;
  WriteLn('WrittenStream');
end.
"#);
    assert_eq!(out, vec!["WrittenStream"]);
}

#[test]
fn test_tfilestream_openread_and_readbuffer() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  fs := TFileStream.Create('test_stream_1.dat', fmCreate);
  val := 500; fs.WriteBuffer(val, SizeOf(Integer));
  fs.Free;

  fs := TFileStream.Create('test_stream_1.dat', fmOpenRead);
  try
    fs.ReadBuffer(val, SizeOf(Integer));
    WriteLn(val);
  finally
    fs.Free;
  end;
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_tfilestream_seek_sofrombeginning() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; v1, v2, val: Integer;
begin
  v1 := 10; v2 := 20;
  fs := TFileStream.Create('test_stream_seek.dat', fmCreate);
  fs.WriteBuffer(v1, SizeOf(Integer));
  fs.WriteBuffer(v2, SizeOf(Integer));

  fs.Seek(0, soFromBeginning);
  fs.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_tfilestream_position_property() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  fs := TFileStream.Create('test_stream_pos.dat', fmCreate);
  val := 100;
  fs.WriteBuffer(val, SizeOf(Integer));
  WriteLn(fs.Position);
  fs.Position := 0;
  WriteLn(fs.Position);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["4", "0"]);
}

#[test]
fn test_tfilestream_size_property() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  fs := TFileStream.Create('test_stream_sz.dat', fmCreate);
  val := 1; fs.WriteBuffer(val, SizeOf(Integer));
  val := 2; fs.WriteBuffer(val, SizeOf(Integer));
  WriteLn(fs.Size);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_tfilestream_seek_sofromend() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  val := 42;
  fs := TFileStream.Create('test_stream_end.dat', fmCreate);
  fs.WriteBuffer(val, SizeOf(Integer));
  fs.Seek(0, soFromEnd);
  WriteLn(fs.Position);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_tfilestream_copyfrom_between_streams() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs1, fs2: TFileStream; val: Integer;
begin
  val := 777;
  fs1 := TFileStream.Create('test_src.dat', fmCreate);
  fs1.WriteBuffer(val, SizeOf(Integer));
  fs1.Position := 0;

  fs2 := TFileStream.Create('test_dst.dat', fmCreate);
  fs2.CopyFrom(fs1, fs1.Size);
  fs2.Position := 0;
  fs2.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);

  fs1.Free; fs2.Free;
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_tfilestream_record_serialization() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TPoint = packed record X, Y: Integer; end;
var fs: TFileStream; pt1, pt2: TPoint;
begin
  pt1.X := 15; pt1.Y := 30;
  fs := TFileStream.Create('test_rec.dat', fmCreate);
  fs.WriteBuffer(pt1, SizeOf(TPoint));
  fs.Position := 0;
  fs.ReadBuffer(pt2, SizeOf(TPoint));
  WriteLn(pt2.X + pt2.Y);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["45"]);
}

#[test]
fn test_tfilestream_string_length_prefixed_serialization() {
    let out = run_pascal(r#"
program Test;
uses Classes;
procedure WriteString(stream: TStream; const s: String);
var len: Integer;
begin
  len := Length(s);
  stream.WriteBuffer(len, SizeOf(Integer));
  if len > 0 then stream.WriteBuffer(s[1], len);
end;
function ReadString(stream: TStream): String;
var len: Integer;
begin
  stream.ReadBuffer(len, SizeOf(Integer));
  SetLength(Result, len);
  if len > 0 then stream.ReadBuffer(Result[1], len);
end;
var fs: TFileStream;
begin
  fs := TFileStream.Create('test_str.dat', fmCreate);
  WriteString(fs, 'StreamedStringData');
  fs.Position := 0;
  WriteLn(ReadString(fs));
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["StreamedStringData"]);
}

#[test]
fn test_tfilestream_polymorphic_parameter_passing() {
    let out = run_pascal(r#"
program Test;
uses Classes;
procedure SaveIntToStream(stream: TStream; val: Integer);
begin
  stream.WriteBuffer(val, SizeOf(Integer));
end;
var fs: TFileStream; v: Integer;
begin
  fs := TFileStream.Create('test_poly.dat', fmCreate);
  SaveIntToStream(fs, 999);
  fs.Position := 0;
  fs.ReadBuffer(v, SizeOf(Integer));
  WriteLn(v);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_tfilestream_openreadwrite_mode() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  fs := TFileStream.Create('test_rw.dat', fmCreate);
  val := 10; fs.WriteBuffer(val, SizeOf(Integer));
  fs.Free;

  fs := TFileStream.Create('test_rw.dat', fmOpenReadWrite);
  fs.Position := 0;
  fs.ReadBuffer(val, SizeOf(Integer));
  val := val + 5;
  fs.Position := 0;
  fs.WriteBuffer(val, SizeOf(Integer));
  fs.Position := 0;
  fs.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_tfilestream_seek_sofromcurrent() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; v1, v2, readVal: Integer;
begin
  v1 := 100; v2 := 200;
  fs := TFileStream.Create('test_cur.dat', fmCreate);
  fs.WriteBuffer(v1, SizeOf(Integer));
  fs.WriteBuffer(v2, SizeOf(Integer));
  fs.Position := 0;
  fs.Seek(4, soFromCurrent);
  fs.ReadBuffer(readVal, SizeOf(Integer));
  WriteLn(readVal);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["200"]);
}

#[test]
fn test_tfilestream_size_truncation() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; val: Integer;
begin
  val := 1;
  fs := TFileStream.Create('test_trunc.dat', fmCreate);
  fs.WriteBuffer(val, SizeOf(Integer));
  fs.WriteBuffer(val, SizeOf(Integer));
  WriteLn(fs.Size);
  fs.Size := 4;
  WriteLn(fs.Size);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["8", "4"]);
}

#[test]
fn test_tfilestream_real_type() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; r, readR: Real;
begin
  r := 45.67;
  fs := TFileStream.Create('test_real.dat', fmCreate);
  fs.WriteBuffer(r, SizeOf(Real));
  fs.Position := 0;
  fs.ReadBuffer(readR, SizeOf(Real));
  WriteLn(readR);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["45.67"]);
}

#[test]
fn test_tfilestream_boolean_type() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream; b, readB: Boolean;
begin
  b := True;
  fs := TFileStream.Create('test_bool.dat', fmCreate);
  fs.WriteBuffer(b, SizeOf(Boolean));
  fs.Position := 0;
  fs.ReadBuffer(readB, SizeOf(Boolean));
  WriteLn(readB);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tfilestream_non_existent_file_raises_efopenerror() {
    let out = run_pascal(r#"
program Test;
uses Classes, SysUtils;
var fs: TFileStream;
begin
  try
    fs := TFileStream.Create('non_existent_file_xyz.dat', fmOpenRead);
    fs.Free;
  except
    on E: EFOpenError do WriteLn('FOpenErrorCaught');
  end;
end.
"#);
    assert_eq!(out, vec!["FOpenErrorCaught"]);
}

#[test]
fn test_tfilestream_byte_array_block_write() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream;
    writeBuf, readBuf: array[0..3] of Byte;
begin
  writeBuf[0] := 1; writeBuf[1] := 2; writeBuf[2] := 3; writeBuf[3] := 4;
  fs := TFileStream.Create('test_bytes.dat', fmCreate);
  fs.WriteBuffer(writeBuf[0], 4);
  fs.Position := 0;
  fs.ReadBuffer(readBuf[0], 4);
  WriteLn(readBuf[0]);
  WriteLn(readBuf[3]);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn test_tfilestream_empty_file_size() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream;
begin
  fs := TFileStream.Create('test_empty_fs.dat', fmCreate);
  WriteLn(fs.Size);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tfilestream_multiple_records() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TItem = packed record ID: Integer; Val: Word; end;
var fs: TFileStream; i1, i2, item: TItem;
begin
  i1.ID := 1; i1.Val := 100;
  i2.ID := 2; i2.Val := 200;
  fs := TFileStream.Create('test_multi_rec.dat', fmCreate);
  fs.WriteBuffer(i1, SizeOf(TItem));
  fs.WriteBuffer(i2, SizeOf(TItem));
  fs.Position := SizeOf(TItem);
  fs.ReadBuffer(item, SizeOf(TItem));
  WriteLn(item.ID.ToString + ':' + item.Val.ToString);
  fs.Free;
end.
"#);
    assert_eq!(out, vec!["2:200"]);
}

#[test]
fn test_tfilestream_protection_in_finally() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var fs: TFileStream;
begin
  fs := TFileStream.Create('test_prot_stream.dat', fmCreate);
  try
    WriteLn('StreamActive');
  finally
    fs.Free;
    WriteLn('StreamFreedInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["StreamActive", "StreamFreedInFinally"]);
}
