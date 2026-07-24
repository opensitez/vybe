use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 22: Dynamic Memory Allocation (GetMem, FreeMem, New, Dispose)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_new_and_dispose_typed_pointer() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  New(p);
  p^ := 100;
  WriteLn(p^);
  Dispose(p);
end.
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_getmem_and_freemem_raw_bytes() {
    let out = run_pascal(r#"
program Test;
var p: Pointer;
    pi: PInteger;
begin
  GetMem(p, SizeOf(Integer));
  pi := PInteger(p);
  pi^ := 500;
  WriteLn(pi^);
  FreeMem(p);
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_new_and_dispose_record_structure() {
    let out = run_pascal(r#"
program Test;
type TNode = record
  ID: Integer;
  Val: String;
end;
type PNode = ^TNode;
var n: PNode;
begin
  New(n);
  n^.ID := 42;
  n^.Val := 'HeapNode';
  WriteLn(n^.ID);
  WriteLn(n^.Val);
  Dispose(n);
end.
"#);
    assert_eq!(out, vec!["42", "HeapNode"]);
}

#[test]
fn test_reallocmem_growing_buffer() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  GetMem(Pointer(p), SizeOf(Integer));
  p^ := 10;
  ReallocMem(Pointer(p), SizeOf(Integer) * 2);
  WriteLn(p^);
  FreeMem(Pointer(p));
end.
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_dynamic_array_setlength_and_length() {
    let out = run_pascal(r#"
program Test;
var arr: array of Integer;
begin
  SetLength(arr, 3);
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  WriteLn(Length(arr));
  WriteLn(arr[1]);
end.
"#);
    assert_eq!(out, vec!["3", "20"]);
}

#[test]
fn test_dynamic_array_high_and_low() {
    let out = run_pascal(r#"
program Test;
var arr: array of String;
begin
  SetLength(arr, 5);
  WriteLn(Low(arr));
  WriteLn(High(arr));
end.
"#);
    assert_eq!(out, vec!["0", "4"]);
}

#[test]
fn test_multidimensional_dynamic_array() {
    let out = run_pascal(r#"
program Test;
var grid: array of array of Integer;
begin
  SetLength(grid, 2, 3);
  grid[1, 2] := 999;
  WriteLn(Length(grid));
  WriteLn(Length(grid[1]));
  WriteLn(grid[1, 2]);
end.
"#);
    assert_eq!(out, vec!["2", "3", "999"]);
}

#[test]
fn test_getmem_with_fillchar_zeroing() {
    let out = run_pascal(r#"
program Test;
var p: PByte;
begin
  GetMem(Pointer(p), 10);
  FillChar(p^, 10, 0);
  WriteLn(p^);
  FreeMem(Pointer(p));
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_dynamic_array_resizing_preserves_elements() {
    let out = run_pascal(r#"
program Test;
var arr: array of Integer;
begin
  SetLength(arr, 2);
  arr[0] := 100; arr[1] := 200;
  SetLength(arr, 4);
  arr[2] := 300; arr[3] := 400;
  WriteLn(arr[0]);
  WriteLn(arr[3]);
end.
"#);
    assert_eq!(out, vec!["100", "400"]);
}

#[test]
fn test_dynamic_array_clearing_via_setlength_zero() {
    let out = run_pascal(r#"
program Test;
var arr: array of Integer;
begin
  SetLength(arr, 5);
  WriteLn(Length(arr));
  SetLength(arr, 0);
  WriteLn(Length(arr));
end.
"#);
    assert_eq!(out, vec!["5", "0"]);
}

#[test]
fn test_new_pointer_in_constructor_and_dispose_in_destructor() {
    let out = run_pascal(r#"
program Test;
type THolder = class
  public PVal: PInteger;
  constructor Create;
  destructor Destroy; override;
end;
constructor THolder.Create; begin New(PVal); PVal^ := 777; end;
destructor THolder.Destroy; begin Dispose(PVal); inherited Destroy; end;
var h: THolder;
begin
  h := THolder.Create;
  WriteLn(h.PVal^);
  h.Free;
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_allocating_array_of_record_pointers() {
    let out = run_pascal(r#"
program Test;
type TRec = record Val: Integer; end;
type PRec = ^TRec;
var ptrs: array[1..2] of PRec;
begin
  New(ptrs[1]); New(ptrs[2]);
  ptrs[1]^.Val := 10; ptrs[2]^.Val := 20;
  WriteLn(ptrs[1]^.Val + ptrs[2]^.Val);
  Dispose(ptrs[1]); Dispose(ptrs[2]);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_dynamic_array_of_strings() {
    let out = run_pascal(r#"
program Test;
var names: array of String;
begin
  SetLength(names, 2);
  names[0] := 'Alice'; names[1] := 'Bob';
  WriteLn(names[0] + ' & ' + names[1]);
end.
"#);
    assert_eq!(out, vec!["Alice & Bob"]);
}

#[test]
fn test_helper_function_allocates_and_returns_pointer() {
    let out = run_pascal(r#"
program Test;
function CreateInt(val: Integer): PInteger;
begin
  New(Result);
  Result^ := val;
end;
var p: PInteger;
begin
  p := CreateInt(888);
  WriteLn(p^);
  Dispose(p);
end.
"#);
    assert_eq!(out, vec!["888"]);
}

#[test]
fn test_reallocmem_shrinking_buffer() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  GetMem(Pointer(p), SizeOf(Integer) * 4);
  p^ := 42;
  ReallocMem(Pointer(p), SizeOf(Integer));
  WriteLn(p^);
  FreeMem(Pointer(p));
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_dynamic_array_copy_function() {
    let out = run_pascal(r#"
program Test;
var src, dst: array of Integer;
begin
  SetLength(src, 3);
  src[0] := 1; src[1] := 2; src[2] := 3;
  dst := Copy(src, 1, 2);
  WriteLn(Length(dst));
  WriteLn(dst[0]);
  WriteLn(dst[1]);
end.
"#);
    assert_eq!(out, vec!["2", "2", "3"]);
}

#[test]
fn test_dynamic_array_concat_function() {
    let out = run_pascal(r#"
program Test;
var a, b, c: array of Integer;
begin
  SetLength(a, 2); a[0] := 10; a[1] := 20;
  SetLength(b, 2); b[0] := 30; b[1] := 40;
  c := Concat(a, b);
  WriteLn(Length(c));
  WriteLn(c[0]);
  WriteLn(c[3]);
end.
"#);
    assert_eq!(out, vec!["4", "10", "40"]);
}

#[test]
fn test_dynamic_array_loop_iteration() {
    let out = run_pascal(r#"
program Test;
var arr: array of Integer;
    i, sum: Integer;
begin
  SetLength(arr, 4);
  arr[0] := 5; arr[1] := 10; arr[2] := 15; arr[3] := 20;
  sum := 0;
  for i := Low(arr) to High(arr) do
    sum := sum + arr[i];
  WriteLn(sum);
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_dynamic_array_in_record_field() {
    let out = run_pascal(r#"
program Test;
type TDataSet = record
  Title: String;
  Values: array of Real;
end;
var ds: TDataSet;
begin
  ds.Title := 'SensorReadings';
  SetLength(ds.Values, 2);
  ds.Values[0] := 12.5; ds.Values[1] := 25.0;
  WriteLn(ds.Title);
  WriteLn(ds.Values[0] + ds.Values[1]);
end.
"#);
    assert_eq!(out, vec!["SensorReadings", "37.5"]);
}

#[test]
fn test_getmem_pointer_not_nil_check() {
    let out = run_pascal(r#"
program Test;
var p: Pointer;
begin
  GetMem(p, 100);
  WriteLn(p <> nil);
  FreeMem(p);
end.
"#);
    assert_eq!(out, vec!["True"]);
}
