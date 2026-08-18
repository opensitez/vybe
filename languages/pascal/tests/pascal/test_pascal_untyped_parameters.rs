use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 28: Untyped Parameters (var Buffer, const Buffer, out Buffer)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_untyped_var_parameter_zeroing() {
    let out = run_pascal(
        r#"
program Test;
procedure ZeroData(var data; size: Integer);
begin
  FillChar(data, size, 0);
end;
var x: Integer;
begin
  x := 999;
  ZeroData(x, SizeOf(Integer));
  WriteLn(x);
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_untyped_const_parameter_reading() {
    let out = run_pascal(
        r#"
program Test;
function SumBytes(const buffer; count: Integer): Integer;
var pb: PByte; i: Integer;
begin
  pb := @buffer; Result := 0;
  for i := 0 to count - 1 do
  begin
    Result := Result + pb^;
    Inc(pb);
  end;
end;
var arr: array[0..2] of Byte;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  WriteLn(SumBytes(arr, 3));
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_untyped_out_parameter_initialization() {
    let out = run_pascal(
        r#"
program Test;
procedure FillPattern(out buffer; count: Integer; pattern: Byte);
begin
  FillByte(buffer, count, pattern);
end;
var bytes: array[0..2] of Byte;
begin
  FillPattern(bytes, 3, $AA);
  WriteLn(bytes[0]);
  WriteLn(bytes[2]);
end.
"#,
    );
    assert_eq!(out, vec!["170", "170"]);
}

#[test]
fn test_untyped_var_parameter_record_mutation() {
    let out = run_pascal(
        r#"
program Test;
type TData = record A, B: Integer; end;
procedure SwapFields(var buffer);
var p1, p2: PInteger; temp: Integer;
begin
  p1 := @buffer;
  p2 := PInteger(PByte(@buffer) + SizeOf(Integer));
  temp := p1^; p1^ := p2^; p2^ := temp;
end;
var d: TData;
begin
  d.A := 10; d.B := 20;
  SwapFields(d);
  WriteLn(d.A);
  WriteLn(d.B);
end.
"#,
    );
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn test_untyped_parameter_move_wrapper() {
    let out = run_pascal(
        r#"
program Test;
procedure CopyBuffer(const src; var dst; size: Integer);
begin
  Move(src, dst, size);
end;
var val1, val2: Integer;
begin
  val1 := 777;
  CopyBuffer(val1, val2, SizeOf(Integer));
  WriteLn(val2);
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_untyped_parameter_with_string_data() {
    let out = run_pascal(
        r#"
program Test;
function FirstCharByte(const buf): Byte;
var pb: PByte;
begin
  pb := @buf;
  Result := pb^;
end;
var strVal: String;
begin
  strVal := 'Pascal';
  WriteLn(FirstCharByte(strVal[1]));
end.
"#,
    );
    assert_eq!(out, vec!["80"]);
}

#[test]
fn test_untyped_parameter_byte_swap_utility() {
    let out = run_pascal(
        r#"
program Test;
procedure ByteSwap16(var buf);
var pb: PByte; temp: Byte;
begin
  pb := @buf;
  temp := pb^;
  pb^ := (pb + 1)^;
  (pb + 1)^ := temp;
end;
var w: Word;
begin
  w := $1234;
  ByteSwap16(w);
  WriteLn(HexStr(w, 4));
end.
"#,
    );
    assert_eq!(out, vec!["3412"]);
}

#[test]
fn test_untyped_parameter_with_float_data() {
    let out = run_pascal(
        r#"
program Test;
procedure DoubleRealBits(var data);
var pr: PReal;
begin
  pr := @data;
  pr^ := pr^ * 2.0;
end;
var r: Real;
begin
  r := 12.5;
  DoubleRealBits(r);
  WriteLn(r);
end.
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_untyped_parameter_with_boolean_array() {
    let out = run_pascal(
        r#"
program Test;
procedure InvertBooleans(var buf; count: Integer);
var pb: PBoolean; i: Integer;
begin
  pb := @buf;
  for i := 1 to count do
  begin
    pb^ := not pb^;
    Inc(pb);
  end;
end;
var flags: array[1..3] of Boolean;
begin
  flags[1] := True; flags[2] := False; flags[3] := True;
  InvertBooleans(flags, 3);
  WriteLn(flags[1]);
  WriteLn(flags[2]);
  WriteLn(flags[3]);
end.
"#,
    );
    assert_eq!(out, vec!["False", "True", "False"]);
}

#[test]
fn test_untyped_parameter_in_class_method() {
    let out = run_pascal(
        r#"
program Test;
type TStreamer = class
  public procedure WriteRaw(const data; size: Integer);
end;
procedure TStreamer.WriteRaw(const data; size: Integer);
var pb: PByte;
begin
  pb := @data;
  WriteLn(pb^);
end;
var s: TStreamer; num: Integer;
begin
  num := 42;
  s := TStreamer.Create;
  s.WriteRaw(num, SizeOf(Integer));
  s.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_untyped_parameter_checksum_calculation() {
    let out = run_pascal(
        r#"
program Test;
function Checksum8(const buf; count: Integer): Byte;
var pb: PByte; i: Integer;
begin
  pb := @buf; Result := 0;
  for i := 0 to count - 1 do
  begin
    Result := Result xor pb^;
    Inc(pb);
  end;
end;
var arr: array[0..2] of Byte;
begin
  arr[0] := $AA; arr[1] := $55; arr[2] := $FF;
  WriteLn(Checksum8(arr, 3));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_untyped_parameter_with_subrange_variable() {
    let out = run_pascal(
        r#"
program Test;
type TSub = 1..100;
procedure InspectSub(const buffer);
var ps: ^TSub;
begin
  ps := @buffer;
  WriteLn(ps^);
end;
var val: TSub;
begin
  val := 75;
  InspectSub(val);
end.
"#,
    );
    assert_eq!(out, vec!["75"]);
}

#[test]
fn test_untyped_parameter_enum_variable() {
    let out = run_pascal(
        r#"
program Test;
type TMode = (mInit, mRun, mStop);
procedure SetModeRaw(var buffer; modeVal: Byte);
var pb: PByte;
begin
  pb := @buffer;
  pb^ := modeVal;
end;
var m: TMode;
begin
  m := mInit;
  SetModeRaw(m, Ord(mRun));
  WriteLn(Ord(m));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_untyped_parameter_multidimensional_array() {
    let out = run_pascal(
        r#"
program Test;
procedure ZeroMatrix(var matrix; elementCount: Integer);
begin
  FillChar(matrix, elementCount * SizeOf(Integer), 0);
end;
var mat: array[0..1, 0..1] of Integer;
begin
  mat[0, 0] := 5; mat[1, 1] := 10;
  ZeroMatrix(mat, 4);
  WriteLn(mat[0, 0]);
  WriteLn(mat[1, 1]);
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn test_untyped_parameter_default_size_wrapper() {
    let out = run_pascal(
        r#"
program Test;
procedure ProcessRaw(var buffer; size: Integer = 4);
var pi: PInteger;
begin
  if size = 4 then
  begin
    pi := @buffer;
    pi^ := pi^ + 10;
  end;
end;
var x: Integer;
begin
  x := 50;
  ProcessRaw(x);
  WriteLn(x);
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_untyped_parameter_record_method() {
    let out = run_pascal(
        r#"
program Test;
type TBufferWrapper = record
  procedure LoadRaw(const source; count: Integer);
end;
procedure TBufferWrapper.LoadRaw(const source; count: Integer);
var pb: PByte;
begin
  pb := @source;
  WriteLn(pb^);
end;
var bw: TBufferWrapper; val: Integer;
begin
  val := 99;
  bw.LoadRaw(val, SizeOf(Integer));
end.
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_untyped_parameter_bitwise_not_buffer() {
    let out = run_pascal(
        r#"
program Test;
procedure InvertBuffer(var buf; size: Integer);
var pb: PByte; i: Integer;
begin
  pb := @buf;
  for i := 1 to size do
  begin
    pb^ := not pb^;
    Inc(pb);
  end;
end;
var b: Byte;
begin
  b := $0F;
  InvertBuffer(b, 1);
  WriteLn(HexStr(b, 2));
end.
"#,
    );
    assert_eq!(out, vec!["F0"]);
}

#[test]
fn test_untyped_parameter_pointer_variable() {
    let out = run_pascal(
        r#"
program Test;
procedure ClearPointerVar(var ptrVar);
var pp: ^Pointer;
begin
  pp := @ptrVar;
  pp^ := nil;
end;
var p: Pointer;
begin
  p := Pointer(12345);
  ClearPointerVar(p);
  WriteLn(p = nil);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_untyped_parameter_shortstring_header() {
    let out = run_pascal(
        r#"
program Test;
function GetShortStringLength(const s): Byte;
var pb: PByte;
begin
  pb := @s;
  Result := pb^;
end;
var ss: ShortString;
begin
  ss := 'Short';
  WriteLn(GetShortStringLength(ss));
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_untyped_parameter_chained_delegation() {
    let out = run_pascal(
        r#"
program Test;
procedure InnerRaw(var buf; sz: Integer);
begin
  FillChar(buf, sz, $FF);
end;
procedure OuterRaw(var buf; sz: Integer);
begin
  InnerRaw(buf, sz);
end;
var b: Byte;
begin
  b := 0;
  OuterRaw(b, 1);
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["255"]);
}
