use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 30: Structure Packing, Alignment & Layout Sizing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_packed_record_sizeof_exact() {
    let out = run_pascal(r#"
program Test;
type TPackedHeader = packed record
  Tag: Byte;
  Code: Word;
  Value: Integer;
end;
begin
  WriteLn(SizeOf(TPackedHeader));
end.
"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_packrecords_directive_one() {
    let out = run_pascal(r#"
program Test;
{$PACKRECORDS 1}
type TRec1 = record
  B: Byte;
  I: Integer;
end;
begin
  WriteLn(SizeOf(TRec1));
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_packed_record_field_offsets() {
    let out = run_pascal(r#"
program Test;
type TPackedRec = packed record
  B1: Byte;
  I1: Integer;
  B2: Byte;
end;
var rec: TPackedRec;
    off1, off2: NativeInt;
begin
  off1 := NativeInt(@rec.I1) - NativeInt(@rec.B1);
  off2 := NativeInt(@rec.B2) - NativeInt(@rec.I1);
  WriteLn(off1);
  WriteLn(off2);
end.
"#);
    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn test_packed_record_binary_deserialization_with_move() {
    let out = run_pascal(r#"
program Test;
type TPacket = packed record
  Id: Byte;
  Length: Word;
  Data: Integer;
end;
var raw: array[0..6] of Byte;
    pkt: TPacket;
begin
  raw[0] := 1;
  raw[1] := 10; raw[2] := 0;
  raw[3] := 42; raw[4] := 0; raw[5] := 0; raw[6] := 0;
  Move(raw[0], pkt, SizeOf(TPacket));
  WriteLn(pkt.Id);
  WriteLn(pkt.Length);
  WriteLn(pkt.Data);
end.
"#);
    assert_eq!(out, vec!["1", "10", "42"]);
}

#[test]
fn test_packed_record_serialization_with_move() {
    let out = run_pascal(r#"
program Test;
type TPacket = packed record
  Cmd: Byte;
  Val: Word;
end;
var pkt: TPacket;
    raw: array[0..2] of Byte;
begin
  pkt.Cmd := 9; pkt.Val := 500;
  Move(pkt, raw[0], SizeOf(TPacket));
  WriteLn(raw[0]);
  WriteLn(raw[1] or (raw[2] shl 8));
end.
"#);
    assert_eq!(out, vec!["9", "500"]);
}

#[test]
fn test_nested_packed_records() {
    let out = run_pascal(r#"
program Test;
type TInner = packed record
  X, Y: Byte;
end;
type TOuter = packed record
  Header: Byte;
  Inner: TInner;
  Footer: Byte;
end;
begin
  WriteLn(SizeOf(TOuter));
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_packed_record_array_contiguity() {
    let out = run_pascal(r#"
program Test;
type TItem = packed record
  Id: Byte;
  Code: Word;
end;
var items: array[0..1] of TItem;
    diff: NativeInt;
begin
  diff := NativeInt(@items[1]) - NativeInt(@items[0]);
  WriteLn(diff);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_packrecords_directive_two() {
    let out = run_pascal(r#"
program Test;
{$PACKRECORDS 2}
type TRec2 = record
  B: Byte;
  W: Word;
end;
begin
  WriteLn(SizeOf(TRec2));
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_packrecords_directive_four() {
    let out = run_pascal(r#"
program Test;
{$PACKRECORDS 4}
type TRec4 = record
  B: Byte;
  I: Integer;
end;
begin
  WriteLn(SizeOf(TRec4));
end.
"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_pointer_to_packed_record() {
    let out = run_pascal(r#"
program Test;
type TPackedData = packed record
  Code: Byte;
  Val: Integer;
end;
type PPackedData = ^TPackedData;
var data: TPackedData;
    p: PPackedData;
begin
  data.Code := 7; data.Val := 777;
  p := @data;
  WriteLn(p^.Code);
  WriteLn(p^.Val);
end.
"#);
    assert_eq!(out, vec!["7", "777"]);
}

#[test]
fn test_packed_record_enum_fields() {
    let out = run_pascal(r#"
program Test;
type TStatus = (stOff, stOn);
type TPackedStatus = packed record
  Flag: Byte;
  Status: TStatus;
end;
var ps: TPackedStatus;
begin
  ps.Flag := 1; ps.Status := stOn;
  WriteLn(SizeOf(TPackedStatus));
  WriteLn(Ord(ps.Status));
end.
"#);
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_packed_record_subrange_fields() {
    let out = run_pascal(r#"
program Test;
type TSmallSub = 1..10;
type TPackedSub = packed record
  Sub: TSmallSub;
  Val: Byte;
end;
var ps: TPackedSub;
begin
  ps.Sub := 5; ps.Val := 100;
  WriteLn(SizeOf(TPackedSub));
  WriteLn(ps.Sub);
end.
"#);
    assert_eq!(out, vec!["2", "5"]);
}

#[test]
fn test_packed_array_byte_elements() {
    let out = run_pascal(r#"
program Test;
type TPackedBytes = packed array[1..5] of Byte;
begin
  WriteLn(SizeOf(TPackedBytes));
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_packed_record_boolean_fields() {
    let out = run_pascal(r#"
program Test;
type TPackedFlags = packed record
  F1: Boolean;
  F2: Boolean;
  F3: Boolean;
end;
var pf: TPackedFlags;
begin
  pf.F1 := True; pf.F2 := False; pf.F3 := True;
  WriteLn(SizeOf(TPackedFlags));
  WriteLn(pf.F1);
  WriteLn(pf.F2);
  WriteLn(pf.F3);
end.
"#);
    assert_eq!(out, vec!["3", "True", "False", "True"]);
}

#[test]
fn test_unpacked_vs_packed_size_comparison() {
    let out = run_pascal(r#"
program Test;
type TUnpacked = record
  B1: Byte;
  I: Integer;
  B2: Byte;
end;
type TPacked = packed record
  B1: Byte;
  I: Integer;
  B2: Byte;
end;
begin
  WriteLn(SizeOf(TPacked) < SizeOf(TUnpacked));
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_packed_record_passed_to_untyped_proc() {
    let out = run_pascal(r#"
program Test;
type TPackedHeader = packed record
  Magic: Word;
  Size: Word;
end;
procedure InspectHeader(const buf; expectedLen: Integer);
var ph: ^TPackedHeader;
begin
  ph := @buf;
  WriteLn(ph^.Magic);
  WriteLn(ph^.Size);
end;
var h: TPackedHeader;
begin
  h.Magic := $4D5A; h.Size := 512;
  InspectHeader(h, SizeOf(TPackedHeader));
end.
"#);
    assert_eq!(out, vec!["19802", "512"]);
}

#[test]
fn test_packed_record_heap_allocation_new() {
    let out = run_pascal(r#"
program Test;
type TPackedItem = packed record
  Id: Byte;
  Code: Word;
end;
type PPackedItem = ^TPackedItem;
var p: PPackedItem;
begin
  New(p);
  p^.Id := 10; p^.Code := 2000;
  WriteLn(p^.Id);
  WriteLn(p^.Code);
  Dispose(p);
end.
"#);
    assert_eq!(out, vec!["10", "2000"]);
}

#[test]
fn test_packrecords_default_restoration() {
    let out = run_pascal(r#"
program Test;
{$PACKRECORDS 1}
type TPacked1 = record B: Byte; I: Integer; end;
{$PACKRECORDS DEFAULT}
type TDefault = record B: Byte; I: Integer; end;
begin
  WriteLn(SizeOf(TPacked1) < SizeOf(TDefault));
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_packed_record_fillchar_initialization() {
    let out = run_pascal(r#"
program Test;
type TPackedData = packed record
  B: Byte;
  W: Word;
  I: Integer;
end;
var d: TPackedData;
begin
  FillChar(d, SizeOf(TPackedData), 0);
  WriteLn(d.B);
  WriteLn(d.W);
  WriteLn(d.I);
end.
"#);
    assert_eq!(out, vec!["0", "0", "0"]);
}

#[test]
fn test_packed_record_int64_alignment() {
    let out = run_pascal(r#"
program Test;
type TPackedInt64 = packed record
  B: Byte;
  V: Int64;
end;
begin
  WriteLn(SizeOf(TPackedInt64));
end.
"#);
    assert_eq!(out, vec!["9"]);
}
