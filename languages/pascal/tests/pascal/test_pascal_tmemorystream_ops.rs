use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 65: Memory Streams (TMemoryStream & RAM Buffers)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tmemorystream_write_read() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; val: Integer;
begin
  ms := TMemoryStream.Create;
  val := 42;
  ms.WriteBuffer(val, SizeOf(Integer));
  ms.Position := 0;
  ms.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_tmemorystream_memory_pointer_access() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; pi: PInteger; val: Integer;
begin
  ms := TMemoryStream.Create;
  val := 999;
  ms.WriteBuffer(val, SizeOf(Integer));
  pi := PInteger(ms.Memory);
  WriteLn(pi^);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_tmemorystream_clear_resets_buffer() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; val: Integer;
begin
  ms := TMemoryStream.Create;
  val := 10; ms.WriteBuffer(val, SizeOf(Integer));
  ms.Clear;
  WriteLn(ms.Size);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tmemorystream_setsize_preallocation() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  ms.SetSize(1024);
  WriteLn(ms.Size);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn test_tmemorystream_savetofile_and_loadfromfile() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms1, ms2: TMemoryStream; val: Integer;
begin
  ms1 := TMemoryStream.Create;
  val := 1234; ms1.WriteBuffer(val, SizeOf(Integer));
  ms1.SaveToFile('test_ms.dat');
  ms1.Free;

  ms2 := TMemoryStream.Create;
  ms2.LoadFromFile('test_ms.dat');
  ms2.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);
  ms2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn test_tmemorystream_savetostream_and_loadfromstream() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms1, ms2: TMemoryStream; val: Integer;
begin
  ms1 := TMemoryStream.Create;
  val := 888; ms1.WriteBuffer(val, SizeOf(Integer));
  ms1.Position := 0;

  ms2 := TMemoryStream.Create;
  ms2.LoadFromStream(ms1);
  ms2.Position := 0;
  ms2.ReadBuffer(val, SizeOf(Integer));
  WriteLn(val);

  ms1.Free; ms2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["888"]);
}

#[test]
fn test_tmemorystream_capacity_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  ms.Capacity := 2048;
  WriteLn(ms.Capacity >= 2048);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tmemorystream_record_struct() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
type TRec = packed record Code: Integer; Val: Word; end;
var ms: TMemoryStream; r1, r2: TRec;
begin
  r1.Code := 50; r1.Val := 1000;
  ms := TMemoryStream.Create;
  ms.WriteBuffer(r1, SizeOf(TRec));
  ms.Position := 0;
  ms.ReadBuffer(r2, SizeOf(TRec));
  WriteLn(r2.Code.ToString + ':' + r2.Val.ToString);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50:1000"]);
}

#[test]
fn test_tmemorystream_string_serialization() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; text, resText: String; len: Integer;
begin
  text := 'MemoryStreamString';
  len := Length(text);
  ms := TMemoryStream.Create;
  ms.WriteBuffer(len, SizeOf(Integer));
  ms.WriteBuffer(text[1], len);

  ms.Position := 0;
  ms.ReadBuffer(len, SizeOf(Integer));
  SetLength(resText, len);
  ms.ReadBuffer(resText[1], len);
  WriteLn(resText);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["MemoryStreamString"]);
}

#[test]
fn test_tmemorystream_seek_operations() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; v1, v2, readVal: Integer;
begin
  v1 := 10; v2 := 20;
  ms := TMemoryStream.Create;
  ms.WriteBuffer(v1, SizeOf(Integer));
  ms.WriteBuffer(v2, SizeOf(Integer));

  ms.Seek(SizeOf(Integer), soFromBeginning);
  ms.ReadBuffer(readVal, SizeOf(Integer));
  WriteLn(readVal);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_tmemorystream_direct_pbyte_iteration() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; pb: PByte; i, sum: Integer;
begin
  ms := TMemoryStream.Create;
  ms.SetSize(3);
  pb := PByte(ms.Memory);
  pb^ := 10; (pb + 1)^ := 20; (pb + 2)^ := 30;

  sum := 0;
  for i := 0 to 2 do
    sum := sum + (pb + i)^;
  WriteLn(sum);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_tmemorystream_protection_finally() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  try
    ms.SetSize(100);
    WriteLn('MemoryStreamActive');
  finally
    ms.Free;
    WriteLn('MemoryStreamFreed');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["MemoryStreamActive", "MemoryStreamFreed"]);
}

#[test]
fn test_tmemorystream_real_values() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; r, readR: Real;
begin
  r := 123.45;
  ms := TMemoryStream.Create;
  ms.WriteBuffer(r, SizeOf(Real));
  ms.Position := 0;
  ms.ReadBuffer(readR, SizeOf(Real));
  WriteLn(readR);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["123.45"]);
}

#[test]
fn test_tmemorystream_boolean_flags() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; b, readB: Boolean;
begin
  b := True;
  ms := TMemoryStream.Create;
  ms.WriteBuffer(b, SizeOf(Boolean));
  ms.Position := 0;
  ms.ReadBuffer(readB, SizeOf(Boolean));
  WriteLn(readB);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tmemorystream_copyfrom_partial() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var src, dst: TMemoryStream; v1, v2, readVal: Integer;
begin
  v1 := 111; v2 := 222;
  src := TMemoryStream.Create;
  src.WriteBuffer(v1, SizeOf(Integer));
  src.WriteBuffer(v2, SizeOf(Integer));
  src.Position := SizeOf(Integer);

  dst := TMemoryStream.Create;
  dst.CopyFrom(src, SizeOf(Integer));
  dst.Position := 0;
  dst.ReadBuffer(readVal, SizeOf(Integer));
  WriteLn(readVal);

  src.Free; dst.Free;
end.
"#,
    );
    assert_eq!(out, vec!["222"]);
}

#[test]
fn test_tmemorystream_polymorphic_helper() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
procedure LogStreamSize(stream: TStream);
begin
  WriteLn(stream.Size);
end;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  ms.SetSize(256);
  LogStreamSize(ms);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["256"]);
}

#[test]
fn test_tmemorystream_array_of_records() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
type TPoint = packed record X, Y: Integer; end;
var ms: TMemoryStream; pts1, pts2: array[0..1] of TPoint;
begin
  pts1[0].X := 1; pts1[0].Y := 2;
  pts1[1].X := 3; pts1[1].Y := 4;
  ms := TMemoryStream.Create;
  ms.WriteBuffer(pts1[0], SizeOf(TPoint) * 2);
  ms.Position := 0;
  ms.ReadBuffer(pts2[0], SizeOf(TPoint) * 2);
  WriteLn(pts2[1].X);
  WriteLn(pts2[1].Y);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn test_tmemorystream_position_truncation() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  ms.SetSize(100);
  ms.Position := 50;
  WriteLn(ms.Position);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_tmemorystream_empty_stream_read() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream;
begin
  ms := TMemoryStream.Create;
  WriteLn(ms.Size = 0);
  WriteLn(ms.Position = 0);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_tmemorystream_loop_appending() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var ms: TMemoryStream; i, val: Integer; sum: Integer;
begin
  ms := TMemoryStream.Create;
  for i := 1 to 3 do
  begin
    val := i * 10;
    ms.WriteBuffer(val, SizeOf(Integer));
  end;

  ms.Position := 0;
  sum := 0;
  while ms.Position < ms.Size do
  begin
    ms.ReadBuffer(val, SizeOf(Integer));
    sum := sum + val;
  end;
  WriteLn(sum);
  ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}
