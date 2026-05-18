/// Tests for complex boolean logic and patterns in Pascal/Delphi:
/// De Morgan's laws, boolean accumulators in loops, truth tables,
/// short-circuit patterns, boolean functions as conditions.

use super::helpers::run_pascal;

// ===================================================================
// DE MORGAN'S LAWS
// ===================================================================

#[test] fn demorgan_not_and() {
    assert_eq!(run_pascal(r#"program T;
var a, b: Boolean;
begin
  a := True;
  b := False;
  WriteLn(not (a and b));
  WriteLn((not a) or (not b));
end."#), &["true", "true"]);
}

#[test] fn demorgan_not_or() {
    assert_eq!(run_pascal(r#"program T;
var a, b: Boolean;
begin
  a := False;
  b := False;
  WriteLn(not (a or b));
  WriteLn((not a) and (not b));
end."#), &["true", "true"]);
}

#[test] fn demorgan_all_cases() {
    assert_eq!(run_pascal(r#"program T;
var a, b: Boolean;
begin
  a := True; b := True;
  WriteLn(not (a and b) = ((not a) or (not b)));
  a := True; b := False;
  WriteLn(not (a and b) = ((not a) or (not b)));
  a := False; b := True;
  WriteLn(not (a and b) = ((not a) or (not b)));
  a := False; b := False;
  WriteLn(not (a and b) = ((not a) or (not b)));
end."#), &["true", "true", "true", "true"]);
}

// ===================================================================
// BOOLEAN ACCUMULATION
// ===================================================================

#[test] fn all_positive_check() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    allPos: Boolean;
    i: Integer;
begin
  SetLength(arr, 4);
  arr[0] := 1; arr[1] := 2; arr[2] := 3; arr[3] := 4;
  allPos := True;
  for i := 0 to 3 do
    if arr[i] <= 0 then allPos := False;
  WriteLn(allPos);
  arr[2] := -1;
  allPos := True;
  for i := 0 to 3 do
    if arr[i] <= 0 then allPos := False;
  WriteLn(allPos);
end."#), &["true", "false"]);
}

#[test] fn any_negative_check() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    anyNeg: Boolean;
    i: Integer;
begin
  SetLength(arr, 4);
  arr[0] := 1; arr[1] := 2; arr[2] := -3; arr[3] := 4;
  anyNeg := False;
  for i := 0 to 3 do
    if arr[i] < 0 then anyNeg := True;
  WriteLn(anyNeg);
end."#), &["true"]);
}

#[test] fn boolean_accumulate_with_early_exit() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    found: Boolean;
    i: Integer;
begin
  SetLength(arr, 5);
  arr[0] := 10; arr[1] := 20; arr[2] := 0; arr[3] := 40; arr[4] := 50;
  found := False;
  for i := 0 to 4 do
  begin
    if arr[i] = 0 then
    begin
      found := True;
      Break;
    end;
  end;
  WriteLn(found);
end."#), &["true"]);
}

// ===================================================================
// COMPOUND BOOLEAN EXPRESSIONS
// ===================================================================

#[test] fn three_way_and() {
    assert_eq!(run_pascal(r#"program T;
var x, y, z: Integer;
begin
  x := 5; y := 10; z := 15;
  WriteLn((x < y) and (y < z) and (x < z));
  z := 3;
  WriteLn((x < y) and (y < z) and (x < z));
end."#), &["true", "false"]);
}

#[test] fn nested_boolean_condition() {
    assert_eq!(run_pascal(r#"program T;
function InRectangle(x, y, x1, y1, x2, y2: Integer): Boolean;
begin
  Result := (x >= x1) and (x <= x2) and (y >= y1) and (y <= y2);
end;
begin
  WriteLn(InRectangle(5, 5, 0, 0, 10, 10));
  WriteLn(InRectangle(15, 5, 0, 0, 10, 10));
end."#), &["true", "false"]);
}

// ===================================================================
// BOOLEAN XOR LOGIC
// ===================================================================

#[test] fn boolean_xor_truth_table() {
    assert_eq!(run_pascal(r#"program T;
begin
  WriteLn(False xor False);
  WriteLn(False xor True);
  WriteLn(True xor False);
  WriteLn(True xor True);
end."#), &["false", "true", "true", "false"]);
}

#[test] fn parity_check() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Boolean;
    parity: Boolean;
    i: Integer;
begin
  SetLength(arr, 4);
  arr[0] := True; arr[1] := False; arr[2] := True; arr[3] := True;
  parity := False;
  for i := 0 to 3 do
    parity := parity xor arr[i];
  WriteLn(parity);
end."#), &["true"]);
}

// ===================================================================
// SHORT-CIRCUIT IN LOOPS
// ===================================================================

#[test] fn short_circuit_prevents_nil_deref() {
    assert_eq!(run_pascal(r#"program T;
type
  TNode = class
  public
    Value: Integer;
    Next: TNode;
  end;
var head, curr: TNode;
    count: Integer;
begin
  head := TNode.Create;
  head.Value := 1;
  head.Next := TNode.Create;
  head.Next.Value := 2;
  head.Next.Next := nil;
  count := 0;
  curr := head;
  while Assigned(curr) and (curr.Value > 0) do
  begin
    Inc(count);
    curr := curr.Next;
  end;
  WriteLn(count);
end."#), &["2"]);
}

#[test] fn short_circuit_guard() {
    assert_eq!(run_pascal(r#"program T;
function SafeDiv(a, b: Integer): Integer;
begin
  if (b <> 0) and (a mod b = 0) then
    Result := a div b
  else
    Result := -1;
end;
begin
  WriteLn(SafeDiv(10, 2));
  WriteLn(SafeDiv(10, 3));
  WriteLn(SafeDiv(0, 0));
end."#), &["5", "-1", "-1"]);
}

// ===================================================================
// BOOLEAN FUNCTION COMPOSITION
// ===================================================================

#[test] fn boolean_and_function_results() {
    assert_eq!(run_pascal(r#"program T;
function IsEven(n: Integer): Boolean;
begin
  Result := n mod 2 = 0;
end;
function IsPositive(n: Integer): Boolean;
begin
  Result := n > 0;
end;
function IsEvenAndPositive(n: Integer): Boolean;
begin
  Result := IsEven(n) and IsPositive(n);
end;
begin
  WriteLn(IsEvenAndPositive(4));
  WriteLn(IsEvenAndPositive(-4));
  WriteLn(IsEvenAndPositive(3));
end."#), &["true", "false", "false"]);
}

#[test] fn boolean_or_of_conditions() {
    assert_eq!(run_pascal(r#"program T;
function IsBoundary(n, size: Integer): Boolean;
begin
  Result := (n = 0) or (n = size - 1);
end;
begin
  WriteLn(IsBoundary(0, 10));
  WriteLn(IsBoundary(9, 10));
  WriteLn(IsBoundary(5, 10));
end."#), &["true", "true", "false"]);
}

// ===================================================================
// BOOLEAN ARRAY OPERATIONS
// ===================================================================

#[test] fn count_true_in_bool_array() {
    assert_eq!(run_pascal(r#"program T;
var flags: array of Boolean;
    i, cnt: Integer;
begin
  SetLength(flags, 6);
  flags[0] := True; flags[1] := False; flags[2] := True;
  flags[3] := True; flags[4] := False; flags[5] := True;
  cnt := 0;
  for i := 0 to 5 do
    if flags[i] then Inc(cnt);
  WriteLn(cnt);
end."#), &["4"]);
}

#[test] fn negate_bool_array() {
    assert_eq!(run_pascal(r#"program T;
var flags: array of Boolean;
    i: Integer;
begin
  SetLength(flags, 3);
  flags[0] := True; flags[1] := False; flags[2] := True;
  for i := 0 to 2 do
    flags[i] := not flags[i];
  for i := 0 to 2 do
    WriteLn(flags[i]);
end."#), &["false", "true", "false"]);
}

// ===================================================================
// IMPLICATION PATTERNS
// ===================================================================

#[test] fn logical_implication() {
    assert_eq!(run_pascal(r#"program T;
function Implies(p, q: Boolean): Boolean;
begin
  Result := (not p) or q;
end;
begin
  WriteLn(Implies(True, True));
  WriteLn(Implies(True, False));
  WriteLn(Implies(False, True));
  WriteLn(Implies(False, False));
end."#), &["true", "false", "true", "true"]);
}

// ===================================================================
// BOOLEAN FLAGS PATTERN
// ===================================================================

#[test] fn state_flags_pattern() {
    assert_eq!(run_pascal(r#"program T;
var isRunning, isPaused, isDone: Boolean;
begin
  isRunning := True;
  isPaused := False;
  isDone := False;
  WriteLn(isRunning and not isPaused and not isDone);
  isPaused := True;
  WriteLn(isRunning and not isPaused and not isDone);
  isPaused := False;
  isDone := True;
  WriteLn(isRunning and not isPaused and not isDone);
end."#), &["true", "false", "false"]);
}

// ===================================================================
// BOOLEAN IN RECORD
// ===================================================================

#[test] fn boolean_fields_in_record() {
    assert_eq!(run_pascal(r#"program T;
type
  TPermission = record
    CanRead: Boolean;
    CanWrite: Boolean;
    CanDelete: Boolean;
    function CanModify: Boolean;
    function FullAccess: Boolean;
  end;
function TPermission.CanModify: Boolean;
begin
  Result := CanRead and CanWrite;
end;
function TPermission.FullAccess: Boolean;
begin
  Result := CanRead and CanWrite and CanDelete;
end;
var p: TPermission;
begin
  p.CanRead := True;
  p.CanWrite := True;
  p.CanDelete := False;
  WriteLn(p.CanModify);
  WriteLn(p.FullAccess);
end."#), &["true", "false"]);
}
