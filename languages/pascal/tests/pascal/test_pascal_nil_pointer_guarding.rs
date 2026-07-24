use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 27: Nil Pointer Guarding & Safe Deferral (Assigned, FreeAndNil)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_assigned_function_unassigned() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  p := nil;
  WriteLn(Assigned(p));
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_assigned_function_assigned() {
    let out = run_pascal(r#"
program Test;
var val: Integer;
    p: PInteger;
begin
  val := 42;
  p := @val;
  WriteLn(Assigned(p));
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_safe_nil_check_before_dereference() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  p := nil;
  if p <> nil then WriteLn(p^)
  else WriteLn('NilGuarded');
end.
"#);
    assert_eq!(out, vec!["NilGuarded"]);
}

#[test]
fn test_short_circuit_eval_prevents_nil_crash() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  p := nil;
  if (p <> nil) and (p^ = 100) then WriteLn('Unreachable')
  else WriteLn('ShortCircuitSafe');
end.
"#);
    assert_eq!(out, vec!["ShortCircuitSafe"]);
}

#[test]
fn test_freeandnil_helper_procedure() {
    let out = run_pascal(r#"
program Test;
type TSampleObj = class end;
var obj: TSampleObj;
begin
  obj := TSampleObj.Create;
  WriteLn(Assigned(obj));
  FreeAndNil(obj);
  WriteLn(Assigned(obj));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_free_on_nil_instance() {
    let out = run_pascal(r#"
program Test;
type TItem = class end;
var item: TItem;
begin
  item := nil;
  item.Free;
  WriteLn('NilFreeNoOp');
end.
"#);
    assert_eq!(out, vec!["NilFreeNoOp"]);
}

#[test]
fn test_procedural_variable_nil_guard() {
    let out = run_pascal(r#"
program Test;
type TProc = procedure;
var p: TProc;
begin
  p := nil;
  if Assigned(p) then p()
  else WriteLn('ProcNilGuarded');
end.
"#);
    assert_eq!(out, vec!["ProcNilGuarded"]);
}

#[test]
fn test_method_pointer_nil_guard() {
    let out = run_pascal(r#"
program Test;
type TEvent = procedure of object;
var ev: TEvent;
begin
  ev := nil;
  if Assigned(ev) then ev()
  else WriteLn('EventNilGuarded');
end.
"#);
    assert_eq!(out, vec!["EventNilGuarded"]);
}

#[test]
fn test_interface_nil_guard() {
    let out = run_pascal(r#"
program Test;
type IService = interface
  ['{11111111-0000-0000-0000-000000000000}']
  procedure Serve;
end;
var srv: IService;
begin
  srv := nil;
  if Assigned(srv) then srv.Serve
  else WriteLn('ServiceNilGuarded');
end.
"#);
    assert_eq!(out, vec!["ServiceNilGuarded"]);
}

#[test]
fn test_linked_list_nil_termination_loop() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record
       Val: Integer;
       Next: PNode;
     end;
var n1, n2, n3: PNode;
    curr: PNode;
    sum: Integer;
begin
  New(n1); New(n2); New(n3);
  n1^.Val := 10; n1^.Next := n2;
  n2^.Val := 20; n2^.Next := n3;
  n3^.Val := 30; n3^.Next := nil;
  sum := 0;
  curr := n1;
  while curr <> nil do
  begin
    sum := sum + curr^.Val;
    curr := curr^.Next;
  end;
  WriteLn(sum);
  Dispose(n1); Dispose(n2); Dispose(n3);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_function_returns_nil_pointer_safely() {
    let out = run_pascal(r#"
program Test;
function FindInt(target: Integer): PInteger;
begin
  if target = 100 then Result := @target else Result := nil;
end;
var p: PInteger;
begin
  p := FindInt(50);
  if not Assigned(p) then WriteLn('NotFound');
end.
"#);
    assert_eq!(out, vec!["NotFound"]);
}

#[test]
fn test_safe_dispose_and_nil_pattern() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  New(p);
  p^ := 99;
  if p <> nil then
  begin
    Dispose(p);
    p := nil;
  end;
  WriteLn(p = nil);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pchar_nil_guard() {
    let out = run_pascal(r#"
program Test;
var pc: PChar;
begin
  pc := nil;
  if pc = nil then WriteLn('NilPChar');
end.
"#);
    assert_eq!(out, vec!["NilPChar"]);
}

#[test]
fn test_record_pointer_field_nil_check() {
    let out = run_pascal(r#"
program Test;
type TData = record
  Payload: PInteger;
end;
var d: TData;
begin
  d.Payload := nil;
  if d.Payload = nil then WriteLn('PayloadNil');
end.
"#);
    assert_eq!(out, vec!["PayloadNil"]);
}

#[test]
fn test_class_field_object_nil_check() {
    let out = run_pascal(r#"
program Test;
type TChild = class end;
type TParent = class
  public Child: TChild;
  constructor Create;
end;
constructor TParent.Create; begin Child := nil; end;
var p: TParent;
begin
  p := TParent.Create;
  if not Assigned(p.Child) then WriteLn('NoChild');
  p.Free;
end.
"#);
    assert_eq!(out, vec!["NoChild"]);
}

#[test]
fn test_guard_nil_before_passing_to_procedure() {
    let out = run_pascal(r#"
program Test;
procedure ProcessInt(p: PInteger);
begin
  if p = nil then Exit;
  WriteLn(p^);
end;
var val: Integer;
begin
  ProcessInt(nil);
  val := 777;
  ProcessInt(@val);
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_nil_comparison_in_case_statement_else() {
    let out = run_pascal(r#"
program Test;
var p: PInteger;
begin
  p := nil;
  if p = nil then WriteLn('BranchNil')
  else WriteLn('BranchAssigned');
end.
"#);
    assert_eq!(out, vec!["BranchNil"]);
}

#[test]
fn test_assigned_macro_in_boolean_expression() {
    let out = run_pascal(r#"
program Test;
var p1, p2: PInteger; x: Integer;
begin
  p1 := nil;
  x := 10;
  p2 := @x;
  WriteLn(Assigned(p1) or Assigned(p2));
  WriteLn(Assigned(p1) and Assigned(p2));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_safe_array_of_pointers_iteration() {
    let out = run_pascal(r#"
program Test;
var ptrs: array[1..3] of PInteger;
    i, count: Integer; v: Integer;
begin
  v := 100;
  ptrs[1] := @v; ptrs[2] := nil; ptrs[3] := @v;
  count := 0;
  for i := 1 to 3 do
    if Assigned(ptrs[i]) then Inc(count);
  WriteLn(count);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_dynamic_array_nil_comparison() {
    let out = run_pascal(r#"
program Test;
var arr: array of Integer;
begin
  arr := nil;
  WriteLn(arr = nil);
  SetLength(arr, 1);
  WriteLn(arr = nil);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}
