/// Tests for Pascal data structures: arrays, sets, nested arrays, string arrays.

use super::helpers::run_pascal;

// ===================================================================
// DYNAMIC ARRAYS
// ===================================================================

#[test] fn dyn_array_empty() {
    assert_eq!(run_pascal("program T; var a: array of Integer; begin a := []; WriteLn(Length(a)); end."), &["0"]);
}

#[test] fn dyn_array_single() {
    assert_eq!(run_pascal("program T; var a: array of Integer; begin a := [42]; WriteLn(a[0]); end."), &["42"]);
}

#[test] fn dyn_array_modify_loop() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i: Integer;
begin
  a := [10, 20, 30, 40, 50];
  for i := 0 to High(a) do a[i] := a[i] * 2;
  for i := 0 to High(a) do WriteLn(a[i]);
end."#), &["20", "40", "60", "80", "100"]);
}

#[test] fn dyn_array_sum() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i, s: Integer;
begin
  a := [1, 2, 3, 4, 5];
  s := 0;
  for i := 0 to High(a) do s := s + a[i];
  WriteLn(s);
end."#), &["15"]);
}

#[test] fn dyn_array_of_strings() {
    assert_eq!(run_pascal(r#"program T;
var names: array of String;
begin
  names := ['Alice', 'Bob', 'Charlie'];
  WriteLn(Length(names));
  WriteLn(names[1]);
end."#), &["3", "Bob"]);
}

#[test] fn dyn_array_of_reals() {
    assert_eq!(run_pascal(r#"program T;
var vals: array of Real;
begin
  vals := [1.5, 2.5, 3.5];
  WriteLn(vals[0] + vals[2]);
end."#), &["5"]);
}

#[test] fn dyn_array_forin_sum() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; x, s: Integer;
begin
  a := [10, 20, 30];
  s := 0;
  for x in a do s := s + x;
  WriteLn(s);
end."#), &["60"]);
}

#[test] fn dyn_array_nested_access() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer;
begin
  a := [100, 200, 300];
  WriteLn(a[a[0] div 100]);
end."#), &["200"]);
}

#[test] fn dyn_array_high_low() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer;
begin
  a := [5, 10, 15, 20];
  WriteLn(Low(a));
  WriteLn(High(a));
  WriteLn(a[Low(a)]);
  WriteLn(a[High(a)]);
end."#), &["0", "3", "5", "20"]);
}

#[test] fn dyn_array_bool() {
    assert_eq!(run_pascal(r#"program T;
var flags: array of Boolean;
begin
  flags := [true, false, true];
  if flags[0] then WriteLn('first is true');
  if not flags[1] then WriteLn('second is false');
end."#), &["first is true", "second is false"]);
}

// ===================================================================
// SET LITERALS (used as array literals in the runtime)
// ===================================================================

#[test] fn set_in_expression() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i: Integer;
begin
  a := [3, 1, 4, 1, 5];
  for i in a do
    if i > 3 then WriteLn(i);
end."#), &["4", "5"]);
}

// ===================================================================
// ARRAY OF CLASSES
// ===================================================================

#[test] fn array_of_objects() {
    assert_eq!(run_pascal(r#"program T;
type TItem = class
  public FName: String;
  constructor Create(N: String);
end;
constructor TItem.Create(N: String); begin FName := N; end;
var items: array of TItem; it: TItem;
begin
  items := [TItem.Create('a'), TItem.Create('b'), TItem.Create('c')];
  for it in items do WriteLn(it.FName);
end."#), &["a", "b", "c"]);
}

// ===================================================================
// MULTI-DIMENSIONAL (array of array)
// ===================================================================

#[test] fn array_2d_manual() {
    assert_eq!(run_pascal(r#"program T;
var row1, row2: array of Integer;
begin
  row1 := [1, 2, 3];
  row2 := [4, 5, 6];
  WriteLn(row1[0] + row2[2]);
end."#), &["7"]);
}

// ===================================================================
// ARRAY IN FUNCTION
// ===================================================================

#[test] fn array_param_function() {
    assert_eq!(run_pascal(r#"program T;
function SumArray(a: array of Integer): Integer;
var i, s: Integer;
begin
  s := 0;
  for i := 0 to High(a) do s := s + a[i];
  Result := s;
end;
begin
  WriteLn(SumArray([10, 20, 30]));
end."#), &["60"]);
}

#[test] fn array_return_from_function() {
    assert_eq!(run_pascal(r#"program T;
function MakeArray: array of Integer;
begin
  Result := [1, 2, 3];
end;
var a: array of Integer;
begin
  a := MakeArray;
  WriteLn(a[0]);
  WriteLn(a[2]);
end."#), &["1", "3"]);
}
