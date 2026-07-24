use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 80: Threading, CriticalSections & Atomic Interlocked Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_thread_tthread_subclassing() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TTestThread = class(TThread)
  protected procedure Execute; override;
end;
procedure TTestThread.Execute;
begin
  WriteLn('ThreadExecuted');
end;
var t: TTestThread;
begin
  t := TTestThread.Create(True); // Create suspended
  t.FreeOnTerminate := False;
  t.Start;
  t.WaitFor;
  t.Free;
end.
"#);
    assert_eq!(out, vec!["ThreadExecuted"]);
}

#[test]
fn test_thread_tcriticalsection_enter_leave() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var cs: TCriticalSection;
begin
  cs := TCriticalSection.Create;
  cs.Enter;
  try
    WriteLn('CriticalSectionEntered');
  finally
    cs.Leave;
  end;
  cs.Free;
end.
"#);
    assert_eq!(out, vec!["CriticalSectionEntered"]);
}

#[test]
fn test_thread_tcriticalsection_acquire_release() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var cs: TCriticalSection;
begin
  cs := TCriticalSection.Create;
  cs.Acquire;
  try
    WriteLn('Acquired');
  finally
    cs.Release;
  end;
  cs.Free;
end.
"#);
    assert_eq!(out, vec!["Acquired"]);
}

#[test]
fn test_thread_tinterlocked_increment_decrement() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var counter: Integer;
begin
  counter := 0;
  TInterlocked.Increment(counter);
  WriteLn(counter);
  TInterlocked.Decrement(counter);
  WriteLn(counter);
end.
"#);
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn test_thread_tinterlocked_add() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var val: Integer;
begin
  val := 10;
  TInterlocked.Add(val, 5);
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_thread_tinterlocked_compareexchange() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var target, oldValue: Integer;
begin
  target := 100;
  oldValue := TInterlocked.CompareExchange(target, 200, 100);
  WriteLn(oldValue.ToString + ':' + target.ToString);
end.
"#);
    assert_eq!(out, vec!["100:200"]);
}

#[test]
fn test_thread_tinterlocked_exchange() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var target, oldVal: Integer;
begin
  target := 50;
  oldVal := TInterlocked.Exchange(target, 75);
  WriteLn(oldVal.ToString + '->' + target.ToString);
end.
"#);
    assert_eq!(out, vec!["50->75"]);
}

#[test]
fn test_thread_tevent_signal_reset() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var event: TEvent; res: TWaitResult;
begin
  event := TEvent.Create(nil, True, False, ''); // Manual reset, initially unsignaled
  event.SetEvent;
  res := event.WaitFor(100);
  WriteLn(res = wrSignaled);

  event.ResetEvent;
  res := event.WaitFor(10);
  WriteLn(res = wrTimeout);

  event.Free;
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_thread_terminated_flag_check() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TLoopThread = class(TThread)
  protected procedure Execute; override;
end;
procedure TLoopThread.Execute;
begin
  while not Terminated do
  begin
    Terminate;
  end;
  WriteLn('TerminatedLoopDone');
end;
var t: TLoopThread;
begin
  t := TLoopThread.Create(True);
  t.FreeOnTerminate := False;
  t.Start;
  t.WaitFor;
  t.Free;
end.
"#);
    assert_eq!(out, vec!["TerminatedLoopDone"]);
}

#[test]
fn test_thread_synchronize_invocation() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TSyncThread = class(TThread)
  private procedure MainThreadWork;
  protected procedure Execute; override;
end;
procedure TSyncThread.MainThreadWork;
begin
  WriteLn('SynchronizedWork');
end;
procedure TSyncThread.Execute;
begin
  Synchronize(MainThreadWork);
end;
var t: TSyncThread;
begin
  t := TSyncThread.Create(True);
  t.FreeOnTerminate := False;
  t.Start;
  t.WaitFor;
  t.Free;
end.
"#);
    assert_eq!(out, vec!["SynchronizedWork"]);
}

#[test]
fn test_thread_tspinlock_tryenter() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var spin: TSpinLock;
begin
  spin.Enter;
  try
    WriteLn('SpinLockEntered');
  finally
    spin.Exit;
  end;
end.
"#);
    assert_eq!(out, vec!["SpinLockEntered"]);
}

#[test]
fn test_thread_multiple_threads_synchronization() {
    let out = run_pascal(r#"
program Test;
uses Classes, SyncObjs;
var sharedCounter: Integer; cs: TCriticalSection;

type TWorkerThread = class(TThread)
  protected procedure Execute; override;
end;
procedure TWorkerThread.Execute;
begin
  cs.Enter;
  try
    Inc(sharedCounter);
  finally
    cs.Leave;
  end;
end;

var t1, t2: TWorkerThread;
begin
  sharedCounter := 0;
  cs := TCriticalSection.Create;

  t1 := TWorkerThread.Create(True); t1.FreeOnTerminate := False;
  t2 := TWorkerThread.Create(True); t2.FreeOnTerminate := False;

  t1.Start; t2.Start;
  t1.WaitFor; t2.WaitFor;

  WriteLn(sharedCounter);

  t1.Free; t2.Free; cs.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_thread_tinterlocked_pointer_exchange() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var p1, p2, oldP: Pointer; v1, v2: Integer;
begin
  v1 := 10; v2 := 20;
  p1 := @v1; p2 := @v2;
  oldP := TInterlocked.Exchange(p1, p2);
  WriteLn(PInteger(oldP)^);
  WriteLn(PInteger(p1)^);
end.
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_thread_tinterlocked_double_exchange() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var dTarget, oldD: Double;
begin
  dTarget := 1.5;
  oldD := TInterlocked.Exchange(dTarget, 3.5);
  WriteLn(oldD = 1.5);
  WriteLn(dTarget = 3.5);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_thread_tsemaphore_acquire_release() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var sem: TSemaphore; res: TWaitResult;
begin
  sem := TSemaphore.Create(nil, 1, 1, '');
  res := sem.WaitFor(100);
  WriteLn(res = wrSignaled);
  sem.Release(1);
  sem.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_thread_return_value_property() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TRetThread = class(TThread)
  protected procedure Execute; override;
end;
procedure TRetThread.Execute;
begin
  ReturnValue := 42;
end;
var t: TRetThread;
begin
  t := TRetThread.Create(True);
  t.FreeOnTerminate := False;
  t.Start;
  t.WaitFor;
  WriteLn(t.ReturnValue);
  t.Free;
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_thread_thread_id_access() {
    let out = run_pascal(r#"
program Test;
uses Classes;
var t: TThread;
begin
  t := TThread.CurrentThread;
  WriteLn(t.ThreadID <> 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_thread_priority_level_setting() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TPriorityThread = class(TThread)
  protected procedure Execute; override;
end;
procedure TPriorityThread.Execute; begin end;

var t: TPriorityThread;
begin
  t := TPriorityThread.Create(True);
  t.Priority := tpNormal;
  WriteLn(t.Priority = tpNormal);
  t.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_thread_queue_asynchronous_dispatch() {
    let out = run_pascal(r#"
program Test;
uses Classes;
type TQueueThread = class(TThread)
  private procedure LogQueued;
  protected procedure Execute; override;
end;
procedure TQueueThread.LogQueued;
begin
  WriteLn('QueuedWorkCompleted');
end;
procedure TQueueThread.Execute;
begin
  Queue(LogQueued);
end;
var t: TQueueThread;
begin
  t := TQueueThread.Create(True);
  t.FreeOnTerminate := False;
  t.Start;
  t.WaitFor;
  t.Free;
end.
"#);
    assert_eq!(out, vec!["QueuedWorkCompleted"]);
}

#[test]
fn test_thread_tinterlocked_bit_test_and_set() {
    let out = run_pascal(r#"
program Test;
uses SyncObjs;
var flags: Integer; oldBit: Boolean;
begin
  flags := 0;
  oldBit := TInterlocked.BitTestAndSet(flags, 2);
  WriteLn(oldBit);
  WriteLn(flags = 4);
end.
"#);
    assert_eq!(out, vec!["False", "True"]);
}
