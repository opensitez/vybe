use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 29: Smart Pointers & Interface Lifetime Wrappers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_smart_pointer_basic_auto_cleanup() {
    let out = run_pascal(r#"
program Test;
type TSampleObj = class
  public destructor Destroy; override;
end;
destructor TSampleObj.Destroy;
begin
  WriteLn('AutoCleanedBySmartPointer');
  inherited Destroy;
end;

type ISmartRef = interface
  ['{11112222-3333-4444-5555-666677778888}']
  function GetInstance: TSampleObj;
end;

type TSmartRefImpl = class(TInterfacedObject, ISmartRef)
  private FInstance: TSampleObj;
  public constructor Create(obj: TSampleObj); destructor Destroy; override;
  public function GetInstance: TSampleObj;
end;

constructor TSmartRefImpl.Create(obj: TSampleObj); begin FInstance := obj; end;
destructor TSmartRefImpl.Destroy; begin FInstance.Free; inherited Destroy; end;
function TSmartRefImpl.GetInstance: TSampleObj; begin Result := FInstance; end;

procedure RunScope;
var ref: ISmartRef;
begin
  ref := TSmartRefImpl.Create(TSampleObj.Create);
end;

begin
  RunScope;
end.
"#);
    assert_eq!(out, vec!["AutoCleanedBySmartPointer"]);
}

#[test]
fn test_smart_pointer_value_access() {
    let out = run_pascal(r#"
program Test;
type TData = class
  public Val: Integer;
  constructor Create(V: Integer);
end;
constructor TData.Create(V: Integer); begin Val := V; end;

type IDataHolder = interface
  ['{22223333-4444-5555-6666-777788889999}']
  function GetData: TData;
end;

type TDataHolderImpl = class(TInterfacedObject, IDataHolder)
  private FData: TData;
  public constructor Create(d: TData); destructor Destroy; override;
  public function GetData: TData;
end;
constructor TDataHolderImpl.Create(d: TData); begin FData := d; end;
destructor TDataHolderImpl.Destroy; begin FData.Free; inherited Destroy; end;
function TDataHolderImpl.GetData: TData; begin Result := FData; end;

var holder: IDataHolder;
begin
  holder := TDataHolderImpl.Create(TData.Create(500));
  WriteLn(holder.GetData.Val);
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_smart_pointer_array_ownership() {
    let out = run_pascal(r#"
program Test;
type TItem = class
  public ID: Integer;
  constructor Create(AID: Integer); destructor Destroy; override;
end;
constructor TItem.Create(AID: Integer); begin ID := AID; end;
destructor TItem.Destroy; begin WriteLn('ItemDestroyed:' + ID.ToString); inherited Destroy; end;

type IItemWrapper = interface
  ['{33334444-5555-6666-7777-888899990000}']
end;
type TItemWrapperImpl = class(TInterfacedObject, IItemWrapper)
  private FItem: TItem;
  public constructor Create(i: TItem); destructor Destroy; override;
end;
constructor TItemWrapperImpl.Create(i: TItem); begin FItem := i; end;
destructor TItemWrapperImpl.Destroy; begin FItem.Free; inherited Destroy; end;

procedure ProcessItems;
var items: array[1..2] of IItemWrapper;
begin
  items[1] := TItemWrapperImpl.Create(TItem.Create(1));
  items[2] := TItemWrapperImpl.Create(TItem.Create(2));
end;

begin
  ProcessItems;
end.
"#);
    assert_eq!(out, vec!["ItemDestroyed:1", "ItemDestroyed:2"]);
}

#[test]
fn test_smart_pointer_reset_managed_reference() {
    let out = run_pascal(r#"
program Test;
type TResource = class
  public Tag: String;
  constructor Create(T: String); destructor Destroy; override;
end;
constructor TResource.Create(T: String); begin Tag := T; end;
destructor TResource.Destroy; begin WriteLn('Freed:' + Tag); inherited Destroy; end;

type IResPtr = interface
  ['{44445555-6666-7777-8888-999900001111}']
  procedure Reset(newRes: TResource);
  function Get: TResource;
end;

type TResPtrImpl = class(TInterfacedObject, IResPtr)
  private FRes: TResource;
  public constructor Create(r: TResource); destructor Destroy; override;
  public procedure Reset(newRes: TResource); function Get: TResource;
end;

constructor TResPtrImpl.Create(r: TResource); begin FRes := r; end;
destructor TResPtrImpl.Destroy; begin FRes.Free; inherited Destroy; end;
procedure TResPtrImpl.Reset(newRes: TResource); begin FRes.Free; FRes := newRes; end;
function TResPtrImpl.Get: TResource; begin Result := FRes; end;

var ptr: IResPtr;
begin
  ptr := TResPtrImpl.Create(TResource.Create('Old'));
  ptr.Reset(TResource.Create('New'));
  WriteLn(ptr.Get.Tag);
end.
"#);
    assert_eq!(out, vec!["Freed:Old", "New", "Freed:New"]);
}

#[test]
fn test_smart_pointer_release_ownership() {
    let out = run_pascal(r#"
program Test;
type TWidget = class
  public Name: String;
  constructor Create(N: String);
end;
constructor TWidget.Create(N: String); begin Name := N; end;

type IWidgetHolder = interface
  ['{55556666-7777-8888-9999-000011112222}']
  function Release: TWidget;
end;

type TWidgetHolderImpl = class(TInterfacedObject, IWidgetHolder)
  private FWidget: TWidget;
  public constructor Create(w: TWidget); destructor Destroy; override;
  public function Release: TWidget;
end;

constructor TWidgetHolderImpl.Create(w: TWidget); begin FWidget := w; end;
destructor TWidgetHolderImpl.Destroy; begin if FWidget <> nil then FWidget.Free; inherited Destroy; end;
function TWidgetHolderImpl.Release: TWidget; begin Result := FWidget; FWidget := nil; end;

var holder: IWidgetHolder; w: TWidget;
begin
  holder := TWidgetHolderImpl.Create(TWidget.Create('OwnedWidget'));
  w := holder.Release;
  WriteLn(w.Name);
  w.Free;
end.
"#);
    assert_eq!(out, vec!["OwnedWidget"]);
}

#[test]
fn test_smart_pointer_validity_check() {
    let out = run_pascal(r#"
program Test;
type TDataObj = class end;
type IDataSmart = interface
  ['{66667777-8888-9999-0000-111122223333}']
  function HasValue: Boolean;
end;
type TDataSmartImpl = class(TInterfacedObject, IDataSmart)
  private FObj: TDataObj;
  public constructor Create(o: TDataObj); destructor Destroy; override;
  public function HasValue: Boolean;
end;
constructor TDataSmartImpl.Create(o: TDataObj); begin FObj := o; end;
destructor TDataSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;
function TDataSmartImpl.HasValue: Boolean; begin Result := FObj <> nil; end;

var s1, s2: IDataSmart;
begin
  s1 := TDataSmartImpl.Create(TDataObj.Create);
  s2 := TDataSmartImpl.Create(nil);
  WriteLn(s1.HasValue);
  WriteLn(s2.HasValue);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_smart_pointer_managing_container_instance() {
    let out = run_pascal(r#"
program Test;
type TStringContainer = class
  public Text: String;
  constructor Create(T: String);
end;
constructor TStringContainer.Create(T: String); begin Text := T; end;

type IStrSmart = interface
  ['{77778888-9999-0000-1111-222233334444}']
  function GetContainer: TStringContainer;
end;
type TStrSmartImpl = class(TInterfacedObject, IStrSmart)
  private FCont: TStringContainer;
  public constructor Create(c: TStringContainer); destructor Destroy; override;
  public function GetContainer: TStringContainer;
end;
constructor TStrSmartImpl.Create(c: TStringContainer); begin FCont := c; end;
destructor TStrSmartImpl.Destroy; begin FCont.Free; inherited Destroy; end;
function TStrSmartImpl.GetContainer: TStringContainer; begin Result := FCont; end;

var smart: IStrSmart;
begin
  smart := TStrSmartImpl.Create(TStringContainer.Create('ContainerText'));
  WriteLn(smart.GetContainer.Text);
end.
"#);
    assert_eq!(out, vec!["ContainerText"]);
}

#[test]
fn test_smart_pointer_in_for_loop_scope() {
    let out = run_pascal(r#"
program Test;
type TLoopObj = class
  public Iteration: Integer;
  constructor Create(I: Integer); destructor Destroy; override;
end;
constructor TLoopObj.Create(I: Integer); begin Iteration := I; end;
destructor TLoopObj.Destroy; begin WriteLn('LoopObjFreed:' + Iteration.ToString); inherited Destroy; end;

type ILoopSmart = interface
  ['{88889999-0000-1111-2222-333344445555}']
end;
type TLoopSmartImpl = class(TInterfacedObject, ILoopSmart)
  private FObj: TLoopObj;
  public constructor Create(o: TLoopObj); destructor Destroy; override;
end;
constructor TLoopSmartImpl.Create(o: TLoopObj); begin FObj := o; end;
destructor TLoopSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;

procedure RunLoop;
var i: Integer; s: ILoopSmart;
begin
  for i := 1 to 2 do
  begin
    s := TLoopSmartImpl.Create(TLoopObj.Create(i));
  end;
end;

begin
  RunLoop;
end.
"#);
    assert_eq!(out, vec!["LoopObjFreed:1", "LoopObjFreed:2"]);
}

#[test]
fn test_smart_pointer_shared_ownership_refcount() {
    let out = run_pascal(r#"
program Test;
type TSharedResource = class
  public Name: String;
  constructor Create(N: String); destructor Destroy; override;
end;
constructor TSharedResource.Create(N: String); begin Name := N; end;
destructor TSharedResource.Destroy; begin WriteLn('SharedFreed:' + Name); inherited Destroy; end;

type ISharedSmart = interface
  ['{99990000-1111-2222-3333-444455556666}']
  function GetRes: TSharedResource;
end;
type TSharedSmartImpl = class(TInterfacedObject, ISharedSmart)
  private FRes: TSharedResource;
  public constructor Create(r: TSharedResource); destructor Destroy; override;
  public function GetRes: TSharedResource;
end;
constructor TSharedSmartImpl.Create(r: TSharedResource); begin FRes := r; end;
destructor TSharedSmartImpl.Destroy; begin FRes.Free; inherited Destroy; end;
function TSharedSmartImpl.GetRes: TSharedResource; begin Result := FRes; end;

var ref1, ref2: ISharedSmart;
begin
  ref1 := TSharedSmartImpl.Create(TSharedResource.Create('SharedDoc'));
  ref2 := ref1;
  ref1 := nil;
  WriteLn('Ref1NilCompleted');
  WriteLn(ref2.GetRes.Name);
  ref2 := nil;
  WriteLn('Ref2NilCompleted');
end.
"#);
    assert_eq!(out, vec!["Ref1NilCompleted", "SharedDoc", "SharedFreed:SharedDoc", "Ref2NilCompleted"]);
}

#[test]
fn test_smart_pointer_passed_to_procedure() {
    let out = run_pascal(r#"
program Test;
type TPayload = class
  public Code: Integer;
  constructor Create(C: Integer);
end;
constructor TPayload.Create(C: Integer); begin Code := C; end;

type IPayloadSmart = interface
  ['{00001111-2222-3333-4444-555566667777}']
  function GetPayload: TPayload;
end;
type TPayloadSmartImpl = class(TInterfacedObject, IPayloadSmart)
  private FPayload: TPayload;
  public constructor Create(p: TPayload); destructor Destroy; override;
  public function GetPayload: TPayload;
end;
constructor TPayloadSmartImpl.Create(p: TPayload); begin FPayload := p; end;
destructor TPayloadSmartImpl.Destroy; begin FPayload.Free; inherited Destroy; end;
function TPayloadSmartImpl.GetPayload: TPayload; begin Result := FPayload; end;

procedure ProcessSmart(smart: IPayloadSmart);
begin
  WriteLn(smart.GetPayload.Code);
end;

var sp: IPayloadSmart;
begin
  sp := TPayloadSmartImpl.Create(TPayload.Create(999));
  ProcessSmart(sp);
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_smart_pointer_in_record_field() {
    let out = run_pascal(r#"
program Test;
type TSubItem = class
  public Title: String;
  constructor Create(T: String); destructor Destroy; override;
end;
constructor TSubItem.Create(T: String); begin Title := T; end;
destructor TSubItem.Destroy; begin WriteLn('SubItemFreed:' + Title); inherited Destroy; end;

type ISubSmart = interface
  ['{11112222-0000-1111-2222-333344445555}']
  function GetSub: TSubItem;
end;
type TSubSmartImpl = class(TInterfacedObject, ISubSmart)
  private FSub: TSubItem;
  public constructor Create(s: TSubItem); destructor Destroy; override;
  public function GetSub: TSubItem;
end;
constructor TSubSmartImpl.Create(s: TSubItem); begin FSub := s; end;
destructor TSubSmartImpl.Destroy; begin FSub.Free; inherited Destroy; end;
function TSubSmartImpl.GetSub: TSubItem; begin Result := FSub; end;

type TOuterRec = record
  SmartField: ISubSmart;
end;

procedure RunRecScope;
var rec: TOuterRec;
begin
  rec.SmartField := TSubSmartImpl.Create(TSubItem.Create('RecFieldItem'));
  WriteLn(rec.SmartField.GetSub.Title);
end;

begin
  RunRecScope;
end.
"#);
    assert_eq!(out, vec!["RecFieldItem", "SubItemFreed:RecFieldItem"]);
}

#[test]
fn test_smart_pointer_recursive_call_frames() {
    let out = run_pascal(r#"
program Test;
type TFrameObj = class
  public Depth: Integer;
  constructor Create(D: Integer); destructor Destroy; override;
end;
constructor TFrameObj.Create(D: Integer); begin Depth := D; end;
destructor TFrameObj.Destroy; begin WriteLn('FrameFreed:' + Depth.ToString); inherited Destroy; end;

type IFrameSmart = interface
  ['{22223333-1111-2222-3333-444455556666}']
end;
type TFrameSmartImpl = class(TInterfacedObject, IFrameSmart)
  private FObj: TFrameObj;
  public constructor Create(o: TFrameObj); destructor Destroy; override;
end;
constructor TFrameSmartImpl.Create(o: TFrameObj); begin FObj := o; end;
destructor TFrameSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;

procedure RecursiveScope(depth: Integer);
var smart: IFrameSmart;
begin
  smart := TFrameSmartImpl.Create(TFrameObj.Create(depth));
  if depth > 1 then RecursiveScope(depth - 1);
end;

begin
  RecursiveScope(3);
end.
"#);
    assert_eq!(out, vec!["FrameFreed:1", "FrameFreed:2", "FrameFreed:3"]);
}

#[test]
fn test_smart_pointer_custom_cleanup_closure() {
    let out = run_pascal(r#"
program Test;
type TCleanupProc = procedure;
type ICustomCleanup = interface
  ['{33334444-2222-3333-4444-555566667777}']
end;
type TCustomCleanupImpl = class(TInterfacedObject, ICustomCleanup)
  private FProc: TCleanupProc;
  public constructor Create(p: TCleanupProc); destructor Destroy; override;
end;
constructor TCustomCleanupImpl.Create(p: TCleanupProc); begin FProc := p; end;
destructor TCustomCleanupImpl.Destroy; begin if Assigned(FProc) then FProc(); inherited Destroy; end;

procedure OnCleanup;
begin
  WriteLn('CustomCleanupCallbackExecuted');
end;

procedure TestCustomCleanup;
var c: ICustomCleanup;
begin
  c := TCustomCleanupImpl.Create(OnCleanup);
end;

begin
  TestCustomCleanup;
end.
"#);
    assert_eq!(out, vec!["CustomCleanupCallbackExecuted"]);
}

#[test]
fn test_smart_pointer_reassignment_cleans_old() {
    let out = run_pascal(r#"
program Test;
type TRefObj = class
  public Name: String;
  constructor Create(N: String); destructor Destroy; override;
end;
constructor TRefObj.Create(N: String); begin Name := N; end;
destructor TRefObj.Destroy; begin WriteLn('ReassignedFreed:' + Name); inherited Destroy; end;

type IRefSmart = interface
  ['{44445555-3333-4444-5555-666677778888}']
end;
type TRefSmartImpl = class(TInterfacedObject, IRefSmart)
  private FObj: TRefObj;
  public constructor Create(o: TRefObj); destructor Destroy; override;
end;
constructor TRefSmartImpl.Create(o: TRefObj); begin FObj := o; end;
destructor TRefSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;

var s: IRefSmart;
begin
  s := TRefSmartImpl.Create(TRefObj.Create('First'));
  s := TRefSmartImpl.Create(TRefObj.Create('Second'));
  WriteLn('ReassignmentDone');
end.
"#);
    assert_eq!(out, vec!["ReassignedFreed:First", "ReassignmentDone", "ReassignedFreed:Second"]);
}

#[test]
fn test_smart_pointer_returning_integer_value() {
    let out = run_pascal(r#"
program Test;
type TIntData = class
  public Value: Integer;
  constructor Create(V: Integer);
end;
constructor TIntData.Create(V: Integer); begin Value := V; end;

type IIntSmart = interface
  ['{55556666-4444-5555-6666-777788889999}']
  function GetVal: Integer;
end;
type TIntSmartImpl = class(TInterfacedObject, IIntSmart)
  private FObj: TIntData;
  public constructor Create(o: TIntData); destructor Destroy; override;
  public function GetVal: Integer;
end;
constructor TIntSmartImpl.Create(o: TIntData); begin FObj := o; end;
destructor TIntSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;
function TIntSmartImpl.GetVal: Integer; begin Result := FObj.Value; end;

var s: IIntSmart;
begin
  s := TIntSmartImpl.Create(TIntData.Create(1234));
  WriteLn(s.GetVal);
end.
"#);
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn test_smart_pointer_returning_boolean_value() {
    let out = run_pascal(r#"
program Test;
type TBoolData = class
  public Active: Boolean;
  constructor Create(A: Boolean);
end;
constructor TBoolData.Create(A: Boolean); begin Active := A; end;

type IBoolSmart = interface
  ['{66667777-5555-6666-7777-888899990000}']
  function IsActive: Boolean;
end;
type TBoolSmartImpl = class(TInterfacedObject, IBoolSmart)
  private FObj: TBoolData;
  public constructor Create(o: TBoolData); destructor Destroy; override;
  public function IsActive: Boolean;
end;
constructor TBoolSmartImpl.Create(o: TBoolData); begin FObj := o; end;
destructor TBoolSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;
function TBoolSmartImpl.IsActive: Boolean; begin Result := FObj.Active; end;

var s: IBoolSmart;
begin
  s := TBoolSmartImpl.Create(TBoolData.Create(True));
  WriteLn(s.IsActive);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_smart_pointer_swap_utility() {
    let out = run_pascal(r#"
program Test;
type TSwapObj = class
  public Name: String;
  constructor Create(N: String);
end;
constructor TSwapObj.Create(N: String); begin Name := N; end;

type ISwapSmart = interface
  ['{77778888-6666-7777-8888-999900001111}']
  function GetName: String;
end;
type TSwapSmartImpl = class(TInterfacedObject, ISwapSmart)
  private FObj: TSwapObj;
  public constructor Create(o: TSwapObj); destructor Destroy; override;
  public function GetName: String;
end;
constructor TSwapSmartImpl.Create(o: TSwapObj); begin FObj := o; end;
destructor TSwapSmartImpl.Destroy; begin FObj.Free; inherited Destroy; end;
function TSwapSmartImpl.GetName: String; begin Result := FObj.Name; end;

procedure SwapSmart(var a, b: ISwapSmart);
var temp: ISwapSmart;
begin
  temp := a; a := b; b := temp;
end;

var s1, s2: ISwapSmart;
begin
  s1 := TSwapSmartImpl.Create(TSwapObj.Create('Obj1'));
  s2 := TSwapSmartImpl.Create(TSwapObj.Create('Obj2'));
  SwapSmart(s1, s2);
  WriteLn(s1.GetName);
  WriteLn(s2.GetName);
end.
"#);
    assert_eq!(out, vec!["Obj2", "Obj1"]);
}

#[test]
fn test_smart_pointer_nil_initialization() {
    let out = run_pascal(r#"
program Test;
type INilSmart = interface
  ['{88889999-7777-8888-9999-000011112222}']
end;
var s: INilSmart;
begin
  s := nil;
  WriteLn(s = nil);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_smart_pointer_method_call_on_managed_object() {
    let out = run_pascal(r#"
program Test;
type TWorker = class
  public procedure PerformWork;
end;
procedure TWorker.PerformWork; begin WriteLn('ManagedWorkExecuted'); end;

type IWorkerSmart = interface
  ['{99990000-8888-9999-0000-111122223333}']
  procedure DoWork;
end;
type TWorkerSmartImpl = class(TInterfacedObject, IWorkerSmart)
  private FWorker: TWorker;
  public constructor Create(w: TWorker); destructor Destroy; override;
  public procedure DoWork;
end;
constructor TWorkerSmartImpl.Create(w: TWorker); begin FWorker := w; end;
destructor TWorkerSmartImpl.Destroy; begin FWorker.Free; inherited Destroy; end;
procedure TWorkerSmartImpl.DoWork; begin FWorker.PerformWork; end;

var smart: IWorkerSmart;
begin
  smart := TWorkerSmartImpl.Create(TWorker.Create);
  smart.DoWork;
end.
"#);
    assert_eq!(out, vec!["ManagedWorkExecuted"]);
}

#[test]
fn test_smart_pointer_nested_struct_destruction_order() {
    let out = run_pascal(r#"
program Test;
type TChildObj = class
  constructor Create; destructor Destroy; override;
end;
constructor TChildObj.Create; begin end;
destructor TChildObj.Destroy; begin WriteLn('ChildObjFreed'); inherited Destroy; end;

type TParentObj = class
  public Child: TChildObj;
  constructor Create; destructor Destroy; override;
end;
constructor TParentObj.Create; begin Child := TChildObj.Create; end;
destructor TParentObj.Destroy; begin Child.Free; WriteLn('ParentObjFreed'); inherited Destroy; end;

type IParentSmart = interface
  ['{00001111-9999-0000-1111-222233334444}']
end;
type TParentSmartImpl = class(TInterfacedObject, IParentSmart)
  private FParent: TParentObj;
  public constructor Create(p: TParentObj); destructor Destroy; override;
end;
constructor TParentSmartImpl.Create(p: TParentObj); begin FParent := p; end;
destructor TParentSmartImpl.Destroy; begin FParent.Free; inherited Destroy; end;

procedure RunParentScope;
var ps: IParentSmart;
begin
  ps := TParentSmartImpl.Create(TParentObj.Create);
end;

begin
  RunParentScope;
end.
"#);
    assert_eq!(out, vec!["ChildObjFreed", "ParentObjFreed"]);
}
