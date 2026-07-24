use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 52: Resource Protection & Guarantee (try...finally Blocks)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_try_finally_normal_execution() {
    let out = run_pascal(r#"
program Test;
begin
  try
    WriteLn('ProtectedWork');
  finally
    WriteLn('ResourceFreed');
  end;
end.
"#);
    assert_eq!(out, vec!["ProtectedWork", "ResourceFreed"]);
}

#[test]
fn test_try_finally_execution_on_exception() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('FailInWork');
    finally
      WriteLn('FinallyRunsOnException');
    end;
  except
    on E: Exception do WriteLn('Caught:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["FinallyRunsOnException", "Caught:FailInWork"]);
}

#[test]
fn test_try_finally_object_free_pattern() {
    let out = run_pascal(r#"
program Test;
type TResObj = class
  destructor Destroy; override;
end;
destructor TResObj.Destroy; begin WriteLn('ObjectFreedInFinally'); inherited Destroy; end;

var obj: TResObj;
begin
  obj := TResObj.Create;
  try
    WriteLn('UsingObject');
  finally
    obj.Free;
  end;
end.
"#);
    assert_eq!(out, vec!["UsingObject", "ObjectFreedInFinally"]);
}

#[test]
fn test_try_finally_runs_on_exit() {
    let out = run_pascal(r#"
program Test;
procedure RunProc;
begin
  try
    WriteLn('BeforeExit');
    Exit;
    WriteLn('Unreachable');
  finally
    WriteLn('FinallyRunsOnExit');
  end;
end;
begin
  RunProc;
end.
"#);
    assert_eq!(out, vec!["BeforeExit", "FinallyRunsOnExit"]);
}

#[test]
fn test_try_finally_runs_on_break() {
    let out = run_pascal(r#"
program Test;
var i: Integer;
begin
  for i := 1 to 3 do
  begin
    try
      if i = 2 then Break;
      WriteLn('Iter:' + i.ToString);
    finally
      WriteLn('FinallyIter:' + i.ToString);
    end;
  end;
end.
"#);
    assert_eq!(out, vec!["Iter:1", "FinallyIter:1", "FinallyIter:2"]);
}

#[test]
fn test_try_finally_runs_on_continue() {
    let out = run_pascal(r#"
program Test;
var i: Integer;
begin
  for i := 1 to 2 do
  begin
    try
      if i = 1 then Continue;
      WriteLn('Iter:' + i.ToString);
    finally
      WriteLn('FinallyIter:' + i.ToString);
    end;
  end;
end.
"#);
    assert_eq!(out, vec!["FinallyIter:1", "Iter:2", "FinallyIter:2"]);
}

#[test]
fn test_try_finally_getmem_freemem_protection() {
    let out = run_pascal(r#"
program Test;
var p: Pointer;
begin
  GetMem(p, 100);
  try
    WriteLn('BufferAllocated');
  finally
    FreeMem(p);
    WriteLn('BufferFreed');
  end;
end.
"#);
    assert_eq!(out, vec!["BufferAllocated", "BufferFreed"]);
}

#[test]
fn test_nested_try_finally_execution_order() {
    let out = run_pascal(r#"
program Test;
begin
  try
    try
      WriteLn('InnerWork');
    finally
      WriteLn('InnerFinally');
    end;
  finally
    WriteLn('OuterFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["InnerWork", "InnerFinally", "OuterFinally"]);
}

#[test]
fn test_state_restore_in_finally() {
    let out = run_pascal(r#"
program Test;
var globalState: Boolean;
procedure ModifyState;
var oldState: Boolean;
begin
  oldState := globalState;
  globalState := True;
  try
    WriteLn('StateModified:' + globalState.ToString);
  finally
    globalState := oldState;
    WriteLn('StateRestored:' + globalState.ToString);
  end;
end;
begin
  globalState := False;
  ModifyState;
end.
"#);
    assert_eq!(out, vec!["StateModified:True", "StateRestored:False"]);
}

#[test]
fn test_try_finally_combined_with_try_except() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var objFreed: Boolean;
begin
  objFreed := False;
  try
    try
      raise Exception.Create('ErrorInTry');
    finally
      objFreed := True;
    end;
  except
    on E: Exception do WriteLn('Handled:' + E.Message);
  end;
  WriteLn('ObjFreedStatus:' + objFreed.ToString);
end.
"#);
    assert_eq!(out, vec!["Handled:ErrorInTry", "ObjFreedStatus:True"]);
}

#[test]
fn test_try_finally_stringlist_protection() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  try
    sl.Add('ProtectedLine');
    WriteLn(sl[0]);
  finally
    sl.Free;
    WriteLn('StringListFreed');
  end;
end.
"#);
    assert_eq!(out, vec!["ProtectedLine", "StringListFreed"]);
}

#[test]
fn test_try_finally_lock_unlock_pattern() {
    let out = run_pascal(r#"
program Test;
type TLock = record
  Locked: Boolean;
  procedure Lock; procedure Unlock;
end;
procedure TLock.Lock; begin Locked := True; WriteLn('Locked'); end;
procedure TLock.Unlock; begin Locked := False; WriteLn('Unlocked'); end;

var l: TLock;
begin
  l.Lock;
  try
    WriteLn('CriticalSectionCode');
  finally
    l.Unlock;
  end;
end.
"#);
    assert_eq!(out, vec!["Locked", "CriticalSectionCode", "Unlocked"]);
}

#[test]
fn test_try_finally_multiple_resources() {
    let out = run_pascal(r#"
program Test;
type TResA = class destructor Destroy; override; end;
type TResB = class destructor Destroy; override; end;
destructor TResA.Destroy; begin WriteLn('ResAFreed'); inherited Destroy; end;
destructor TResB.Destroy; begin WriteLn('ResBFreed'); inherited Destroy; end;

var a: TResA; b: TResB;
begin
  a := TResA.Create;
  try
    b := TResB.Create;
    try
      WriteLn('UsingBothResources');
    finally
      b.Free;
    end;
  finally
    a.Free;
  end;
end.
"#);
    assert_eq!(out, vec!["UsingBothResources", "ResBFreed", "ResAFreed"]);
}

#[test]
fn test_try_finally_in_function_return() {
    let out = run_pascal(r#"
program Test;
function ComputeValue: Integer;
begin
  try
    Result := 42;
  finally
    WriteLn('FinallyInFunction');
  end;
end;
begin
  WriteLn(ComputeValue);
end.
"#);
    assert_eq!(out, vec!["FinallyInFunction", "42"]);
}

#[test]
fn test_try_finally_in_recursive_procedure() {
    let out = run_pascal(r#"
program Test;
procedure RecursiveProc(depth: Integer);
begin
  try
    WriteLn('Enter:' + depth.ToString);
    if depth > 1 then RecursiveProc(depth - 1);
  finally
    WriteLn('Leave:' + depth.ToString);
  end;
end;
begin
  RecursiveProc(2);
end.
"#);
    assert_eq!(out, vec!["Enter:2", "Enter:1", "Leave:1", "Leave:2"]);
}

#[test]
fn test_try_finally_cursor_toggle_restoration() {
    let out = run_pascal(r#"
program Test;
var cursor: String;
procedure ShowWaitCursor;
var oldCursor: String;
begin
  oldCursor := cursor;
  cursor := 'Wait';
  try
    WriteLn('CurrentCursor:' + cursor);
  finally
    cursor := oldCursor;
    WriteLn('RestoredCursor:' + cursor);
  end;
end;
begin
  cursor := 'Default';
  ShowWaitCursor;
end.
"#);
    assert_eq!(out, vec!["CurrentCursor:Wait", "RestoredCursor:Default"]);
}

#[test]
fn test_try_finally_unhandled_exception_propagation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure DoWork;
begin
  try
    raise Exception.Create('UncaughtInSub');
  finally
    WriteLn('SubFinallyExecuted');
  end;
end;
begin
  try
    DoWork;
  except
    on E: Exception do WriteLn('TopCaught:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["SubFinallyExecuted", "TopCaught:UncaughtInSub"]);
}

#[test]
fn test_try_finally_nil_check_free_pattern() {
    let out = run_pascal(r#"
program Test;
type TCleanObj = class end;
var obj: TCleanObj;
begin
  obj := nil;
  try
    WriteLn('NoAllocationMade');
  finally
    obj.Free;
    WriteLn('NilFreeHandled');
  end;
end.
"#);
    assert_eq!(out, vec!["NoAllocationMade", "NilFreeHandled"]);
}

#[test]
fn test_try_finally_in_record_method() {
    let out = run_pascal(r#"
program Test;
type TWorkerRec = record
  procedure Execute;
end;
procedure TWorkerRec.Execute;
begin
  try
    WriteLn('RecExecuteStart');
  finally
    WriteLn('RecExecuteFinally');
  end;
end;
var w: TWorkerRec;
begin
  w.Execute;
end.
"#);
    assert_eq!(out, vec!["RecExecuteStart", "RecExecuteFinally"]);
}

#[test]
fn test_try_finally_counter_increment_decrement() {
    let out = run_pascal(r#"
program Test;
var activeTasks: Integer;
procedure TaskRun;
begin
  Inc(activeTasks);
  try
    WriteLn('ActiveTasks:' + activeTasks.ToString);
  finally
    Dec(activeTasks);
  end;
end;
begin
  activeTasks := 0;
  TaskRun;
  WriteLn('ActiveTasks:' + activeTasks.ToString);
end.
"#);
    assert_eq!(out, vec!["ActiveTasks:1", "ActiveTasks:0"]);
}
