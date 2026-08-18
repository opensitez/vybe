use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 26: Custom Memory Manager & Heap Interception
// ═══════════════════════════════════════════════════════════

#[test]
fn test_getmemorymanager_query() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
begin
  GetMemoryManager(mm);
  WriteLn(Assigned(mm.GetMem));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_memory_manager_interception_counters() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    AllocCount, FreeCount: Integer;

function CustomGetMem(Size: NativeInt): Pointer;
begin
  Inc(AllocCount);
  Result := OldMM.GetMem(Size);
end;

function CustomFreeMem(P: Pointer): Integer;
begin
  Inc(FreeCount);
  Result := OldMM.FreeMem(P);
end;

var p: Pointer;
begin
  AllocCount := 0; FreeCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.GetMem := CustomGetMem;
  NewMM.FreeMem := CustomFreeMem;
  SetMemoryManager(NewMM);

  GetMem(p, 100);
  FreeMem(p);

  SetMemoryManager(OldMM);
  WriteLn(AllocCount);
  WriteLn(FreeCount);
end.
"#,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn test_custom_memory_manager_realloc_interception() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    ReallocCount: Integer;

function CustomReallocMem(P: Pointer; Size: NativeInt): Pointer;
begin
  Inc(ReallocCount);
  Result := OldMM.ReallocMem(P, Size);
end;

var p: Pointer;
begin
  ReallocCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.ReallocMem := CustomReallocMem;
  SetMemoryManager(NewMM);

  GetMem(p, 50);
  ReallocMem(p, 100);
  FreeMem(p);

  SetMemoryManager(OldMM);
  WriteLn(ReallocCount);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_memory_manager_swap_restoration() {
    let out = run_pascal(
        r#"
program Test;
var OriginalMM, TempMM, CurrentMM: TMemoryManagerEx;
begin
  GetMemoryManager(OriginalMM);
  TempMM := OriginalMM;
  SetMemoryManager(TempMM);
  GetMemoryManager(CurrentMM);
  WriteLn(Assigned(CurrentMM.GetMem));
  SetMemoryManager(OriginalMM);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_allocated_bytes_tracking() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    TotalAllocated: NativeInt;

function TrackGetMem(Size: NativeInt): Pointer;
begin
  TotalAllocated := TotalAllocated + Size;
  Result := OldMM.GetMem(Size);
end;

var p1, p2: Pointer;
begin
  TotalAllocated := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.GetMem := TrackGetMem;
  SetMemoryManager(NewMM);

  GetMem(p1, 64);
  GetMem(p2, 128);
  FreeMem(p1);
  FreeMem(p2);

  SetMemoryManager(OldMM);
  WriteLn(TotalAllocated);
end.
"#,
    );
    assert_eq!(out, vec!["192"]);
}

#[test]
fn test_custom_mem_manager_alloc_and_free_balance_check() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    ActivePointers: Integer;

function CountGetMem(Size: NativeInt): Pointer;
begin
  Inc(ActivePointers);
  Result := OldMM.GetMem(Size);
end;

function CountFreeMem(P: Pointer): Integer;
begin
  Dec(ActivePointers);
  Result := OldMM.FreeMem(P);
end;

var ptrs: array[1..3] of Pointer; i: Integer;
begin
  ActivePointers := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.GetMem := CountGetMem;
  NewMM.FreeMem := CountFreeMem;
  SetMemoryManager(NewMM);

  for i := 1 to 3 do GetMem(ptrs[i], 32);
  WriteLn(ActivePointers);
  for i := 1 to 3 do FreeMem(ptrs[i]);
  WriteLn(ActivePointers);

  SetMemoryManager(OldMM);
end.
"#,
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn test_custom_mem_manager_register_expected_memory_leak() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    LeakCount: Integer;

function LeakTrackerFreeMem(P: Pointer): Integer;
begin
  Result := OldMM.FreeMem(P);
end;

begin
  LeakCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  SetMemoryManager(NewMM);

  WriteLn('LeakTrackerInitialized');
  SetMemoryManager(OldMM);
end.
"#,
    );
    assert_eq!(out, vec!["LeakTrackerInitialized"]);
}

#[test]
fn test_custom_mem_manager_allocvec_routine() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    AllocVecCount: Integer;

function CustomAllocMem(Size: NativeInt): Pointer;
begin
  Inc(AllocVecCount);
  Result := OldMM.AllocMem(Size);
end;

var p: Pointer;
begin
  AllocVecCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.AllocMem := CustomAllocMem;
  SetMemoryManager(NewMM);

  p := AllocMem(64);
  FreeMem(p);

  SetMemoryManager(OldMM);
  WriteLn(AllocVecCount);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_custom_mem_manager_zero_size_allocation() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
    p: Pointer;
begin
  GetMemoryManager(mm);
  p := mm.GetMem(0);
  WriteLn(p = nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_multiple_swaps() {
    let out = run_pascal(
        r#"
program Test;
var OriginalMM, CustomMM1, CustomMM2: TMemoryManagerEx;
begin
  GetMemoryManager(OriginalMM);
  CustomMM1 := OriginalMM;
  CustomMM2 := OriginalMM;
  SetMemoryManager(CustomMM1);
  SetMemoryManager(CustomMM2);
  SetMemoryManager(OriginalMM);
  WriteLn('SwapsCompleted');
end.
"#,
    );
    assert_eq!(out, vec!["SwappedSuccessfully"]);
}

#[test]
fn test_custom_mem_manager_new_instance_dispatch() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    NewOpCount: Integer;

function NewOpGetMem(Size: NativeInt): Pointer;
begin
  Inc(NewOpCount);
  Result := OldMM.GetMem(Size);
end;

type TData = record Val: Integer; end;
     PData = ^TData;
var p: PData;
begin
  NewOpCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.GetMem := NewOpGetMem;
  SetMemoryManager(NewMM);

  New(p);
  p^.Val := 123;
  WriteLn(p^.Val);
  Dispose(p);

  SetMemoryManager(OldMM);
  WriteLn(NewOpCount > 0);
end.
"#,
    );
    assert_eq!(out, vec!["123", "True"]);
}

#[test]
fn test_custom_mem_manager_freemem_nil_handling() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
    res: Integer;
begin
  GetMemoryManager(mm);
  res := mm.FreeMem(nil);
  WriteLn(res = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_realloc_nil_pointer() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
    p: Pointer;
begin
  GetMemoryManager(mm);
  p := mm.ReallocMem(nil, 64);
  WriteLn(p <> nil);
  mm.FreeMem(p);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_realloc_zero_size() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
    p: Pointer;
begin
  GetMemoryManager(mm);
  p := mm.GetMem(64);
  p := mm.ReallocMem(p, 0);
  WriteLn(p = nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_header_tagging() {
    let out = run_pascal(
        r#"
program Test;
var OldMM, NewMM: TMemoryManagerEx;
    TagCount: Integer;

function TaggingGetMem(Size: NativeInt): Pointer;
begin
  Inc(TagCount);
  Result := OldMM.GetMem(Size);
end;

var p: Pointer;
begin
  TagCount := 0;
  GetMemoryManager(OldMM);
  NewMM := OldMM;
  NewMM.GetMem := TaggingGetMem;
  SetMemoryManager(NewMM);

  GetMem(p, 32);
  FreeMem(p);

  SetMemoryManager(OldMM);
  WriteLn(TagCount);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_custom_mem_manager_nested_scope_allocation() {
    let out = run_pascal(
        r#"
program Test;
procedure RunScope;
var p: Pointer;
begin
  GetMem(p, 128);
  FillChar(p^, 128, $AA);
  FreeMem(p);
end;
begin
  RunScope;
  WriteLn('ScopeCleared');
end.
"#,
    );
    assert_eq!(out, vec!["ScopeCleared"]);
}

#[test]
fn test_custom_mem_manager_class_instance_allocations() {
    let out = run_pascal(
        r#"
program Test;
type TTestObj = class
  public Val: Integer;
end;
var obj: TTestObj;
begin
  obj := TTestObj.Create;
  obj.Val := 999;
  WriteLn(obj.Val);
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_custom_mem_manager_large_buffer_allocation() {
    let out = run_pascal(
        r#"
program Test;
var p: Pointer;
begin
  GetMem(p, 65536);
  WriteLn(p <> nil);
  FreeMem(p);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_custom_mem_manager_is_memory_manager_set() {
    let out = run_pascal(
        r#"
program Test;
var mm: TMemoryManagerEx;
begin
  GetMemoryManager(mm);
  WriteLn(IsMemoryManagerSet);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_custom_mem_manager_allocmem_zero_fill() {
    let out = run_pascal(
        r#"
program Test;
var pb: PByte;
begin
  pb := PByte(AllocMem(10));
  WriteLn(pb^);
  FreeMem(pb);
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}
