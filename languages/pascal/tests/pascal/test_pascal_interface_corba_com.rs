use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 94: CORBA vs COM Interfaces & Reference Counting Model
// ═══════════════════════════════════════════════════════════

#[test]
fn test_interfaces_com_mode_refcounting() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES COM}
type IComWork = interface
  ['{11111111-1111-1111-1111-111111111111}']
  procedure DoComWork;
end;

type TComImpl = class(TInterfacedObject, IComWork)
  public procedure DoComWork;
end;
procedure TComImpl.DoComWork; begin WriteLn('ComWorkExecuted'); end;

var w: IComWork;
begin
  w := TComImpl.Create;
  w.DoComWork;
end.
"#,
    );
    assert_eq!(out, vec!["ComWorkExecuted"]);
}

#[test]
fn test_interfaces_corba_mode_no_refcounting() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type ICorbaWork = interface
  ['{22222222-2222-2222-2222-222222222222}']
  procedure DoCorbaWork;
end;

type TCorbaImpl = class(TObject, ICorbaWork)
  public procedure DoCorbaWork;
end;
procedure TCorbaImpl.DoCorbaWork; begin WriteLn('CorbaWorkExecuted'); end;

var obj: TCorbaImpl; c: ICorbaWork;
begin
  obj := TCorbaImpl.Create;
  c := obj;
  c.DoCorbaWork;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CorbaWorkExecuted"]);
}

#[test]
fn test_interfaces_supports_query_com() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ITargetIntf = interface
  ['{33333333-3333-3333-3333-333333333333}']
  procedure Action;
end;

type TTargetImpl = class(TInterfacedObject, ITargetIntf)
  public procedure Action;
end;
procedure TTargetImpl.Action; begin WriteLn('ActionSupported'); end;

var unk: IUnknown; t: ITargetIntf;
begin
  unk := TTargetImpl.Create;
  if Supports(unk, ITargetIntf, t) then
    t.Action;
end.
"#,
    );
    assert_eq!(out, vec!["ActionSupported"]);
}

#[test]
fn test_interfaces_as_operator_com() {
    let out = run_pascal(
        r#"
program Test;
type IFoo = interface ['{44444444-4444-4444-4444-444444444444}'] procedure Foo; end;
type IBar = interface ['{55555555-5555-5555-5555-555555555555}'] procedure Bar; end;

type TFooBar = class(TInterfacedObject, IFoo, IBar)
  public procedure Foo; procedure Bar;
end;
procedure TFooBar.Foo; begin WriteLn('FooCall'); end;
procedure TFooBar.Bar; begin WriteLn('BarCall'); end;

var f: IFoo; b: IBar;
begin
  f := TFooBar.Create;
  b := f as IBar;
  b.Bar;
end.
"#,
    );
    assert_eq!(out, vec!["BarCall"]);
}

#[test]
fn test_interfaces_corba_manual_lifetime() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type ICorbaItem = interface
  ['{66666666-6666-6666-6666-666666666666}']
  function GetVal: Integer;
end;

type TCorbaItem = class(TObject, ICorbaItem)
  public function GetVal: Integer;
  destructor Destroy; override;
end;
function TCorbaItem.GetVal: Integer; begin Result := 500; end;
destructor TCorbaItem.Destroy; begin WriteLn('CorbaObjectDestroyedManually'); inherited Destroy; end;

var itemObj: TCorbaItem; itemIntf: ICorbaItem;
begin
  itemObj := TCorbaItem.Create;
  itemIntf := itemObj;
  WriteLn(itemIntf.GetVal);
  itemObj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["500", "CorbaObjectDestroyedManually"]);
}

#[test]
fn test_interfaces_iunknown_methods_override() {
    let out = run_pascal(
        r#"
program Test;
type TCustomUnk = class(TObject, IUnknown)
  private FRefCount: Integer;
  public
    function QueryInterface(constref iid: TGUID; out obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
end;
function TCustomUnk.QueryInterface(constref iid: TGUID; out obj): HResult; begin Result := 0; end;
function TCustomUnk._AddRef: Integer; begin Inc(FRefCount); WriteLn('AddRef:' + FRefCount.ToString); Result := FRefCount; end;
function TCustomUnk._Release: Integer; begin Dec(FRefCount); WriteLn('Release:' + FRefCount.ToString); Result := FRefCount; end;

var unk: IUnknown;
begin
  unk := TCustomUnk.Create;
end.
"#,
    );
    assert_eq!(out, vec!["AddRef:1", "Release:0"]);
}

#[test]
fn test_interfaces_corba_multiple_interface_implementation() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type IA = interface ['{77777777-7777-7777-7777-777777777777}'] procedure DoA; end;
type IB = interface ['{88888888-8888-8888-8888-888888888888}'] procedure DoB; end;

type TCorbaDual = class(TObject, IA, IB)
  public procedure DoA; procedure DoB;
end;
procedure TCorbaDual.DoA; begin WriteLn('CorbaA'); end;
procedure TCorbaDual.DoB; begin WriteLn('CorbaB'); end;

var obj: TCorbaDual; a: IA; b: IB;
begin
  obj := TCorbaDual.Create;
  a := obj; b := obj;
  a.DoA; b.DoB;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CorbaA", "CorbaB"]);
}

#[test]
fn test_interfaces_com_nil_assignment_releases() {
    let out = run_pascal(
        r#"
program Test;
type IData = interface ['{99999999-9999-9999-9999-999999999999}'] end;
type TData = class(TInterfacedObject, IData)
  destructor Destroy; override;
end;
destructor TData.Destroy; begin WriteLn('DataReleasedOnNil'); inherited Destroy; end;

var d: IData;
begin
  d := TData.Create;
  d := nil;
  WriteLn('AfterNil');
end.
"#,
    );
    assert_eq!(out, vec!["DataReleasedOnNil", "AfterNil"]);
}

#[test]
fn test_interfaces_corba_is_operator_query() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type ICheck = interface ['{A0A0A0A0-A0A0-A0A0-A0A0-A0A0A0A0A0A0}'] end;
type TCheckImpl = class(TObject, ICheck) end;

var obj: TObject;
begin
  obj := TCheckImpl.Create;
  WriteLn(obj is ICheck);
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interfaces_supports_class_type_check() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type IService = interface ['{B0B0B0B0-B0B0-B0B0-B0B0-B0B0B0B0B0B0}'] end;
type TServiceImpl = class(TInterfacedObject, IService) end;

var obj: TObject;
begin
  obj := TServiceImpl.Create;
  WriteLn(Supports(obj, IService));
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interfaces_com_refcount_property() {
    let out = run_pascal(
        r#"
program Test;
type TRefObj = class(TInterfacedObject)
  public property RefCount: Integer read FRefCount;
end;

var obj: TRefObj; intf: IUnknown;
begin
  obj := TRefObj.Create;
  intf := obj;
  WriteLn(obj.RefCount = 1);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interfaces_corba_inheritance() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type IBaseCorba = interface ['{C0C0C0C0-C0C0-C0C0-C0C0-C0C0C0C0C0C0}'] procedure BaseProc; end;
type ISubCorba = interface(IBaseCorba) ['{D0D0D0D0-D0D0-D0D0-D0D0-D0D0D0D0D0D0}'] procedure SubProc; end;

type TSubImpl = class(TObject, ISubCorba)
  public procedure BaseProc; procedure SubProc;
end;
procedure TSubImpl.BaseProc; begin WriteLn('BaseCorbaExecuted'); end;
procedure TSubImpl.SubProc; begin WriteLn('SubCorbaExecuted'); end;

var obj: TSubImpl; s: ISubCorba;
begin
  obj := TSubImpl.Create;
  s := obj;
  s.BaseProc; s.SubProc;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["BaseCorbaExecuted", "SubCorbaExecuted"]);
}

#[test]
fn test_interfaces_com_procedure_parameter_passing() {
    let out = run_pascal(
        r#"
program Test;
type IWorker = interface ['{E0E0E0E0-E0E0-E0E0-E0E0-E0E0E0E0E0E0}'] procedure Work; end;
type TWorkerImpl = class(TInterfacedObject, IWorker)
  public procedure Work;
end;
procedure TWorkerImpl.Work; begin WriteLn('WorkParameterDone'); end;

procedure ExecuteWorker(w: IWorker);
begin
  w.Work;
end;

begin
  ExecuteWorker(TWorkerImpl.Create);
end.
"#,
    );
    assert_eq!(out, vec!["WorkParameterDone"]);
}

#[test]
fn test_interfaces_corba_array_of_interfaces() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type ITask = interface ['{F0F0F0F0-F0F0-F0F0-F0F0-F0F0F0F0F0F0}'] procedure Exec; end;
type TTaskImpl = class(TObject, ITask)
  private FName: String;
  public constructor Create(const n: String); procedure Exec;
end;
constructor TTaskImpl.Create(const n: String); begin FName := n; end;
procedure TTaskImpl.Exec; begin WriteLn(FName); end;

var o1, o2: TTaskImpl; tasks: array[0..1] of ITask;
begin
  o1 := TTaskImpl.Create('T1'); o2 := TTaskImpl.Create('T2');
  tasks[0] := o1; tasks[1] := o2;
  tasks[0].Exec; tasks[1].Exec;
  o1.Free; o2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["T1", "T2"]);
}

#[test]
fn test_interfaces_guid_tostring_match() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type IGuidTest = interface ['{12345678-1234-1234-1234-123456789012}'] end;
begin
  WriteLn(GUIDToString(IGuidTest) = '{12345678-1234-1234-1234-123456789012}');
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interfaces_com_weak_ref_avoidance() {
    let out = run_pascal(
        r#"
program Test;
type ISimple = interface ['{11223344-5566-7788-9900-AABBCCDDEEFF}'] procedure Run; end;
type TSimple = class(TInterfacedObject, ISimple)
  public procedure Run;
end;
procedure TSimple.Run; begin WriteLn('RunSimple'); end;

procedure TestScope;
var s: ISimple;
begin
  s := TSimple.Create;
  s.Run;
end;
begin
  TestScope;
  WriteLn('ScopeEnded');
end.
"#,
    );
    assert_eq!(out, vec!["RunSimple", "ScopeEnded"]);
}

#[test]
fn test_interfaces_corba_generic_interface() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type IGenericCorba<T> = interface
  ['{11112222-3333-4444-5555-666677778888}']
  function GetValue: T;
end;

type TGenImpl = class(TObject, IGenericCorba<Integer>)
  public function GetValue: Integer;
end;
function TGenImpl.GetValue: Integer; begin Result := 888; end;

var obj: TGenImpl; intf: IGenericCorba<Integer>;
begin
  obj := TGenImpl.Create;
  intf := obj;
  WriteLn(intf.GetValue);
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["888"]);
}

#[test]
fn test_interfaces_com_interface_return_from_function() {
    let out = run_pascal(
        r#"
program Test;
type IItem = interface ['{AA11BB22-CC33-DD44-EE55-FF6677889900}'] function Name: String; end;
type TItemImpl = class(TInterfacedObject, IItem)
  public function Name: String;
end;
function TItemImpl.Name: String; begin Result := 'ItemReturnedFromFunc'; end;

function CreateItem: IItem;
begin
  Result := TItemImpl.Create;
end;

begin
  WriteLn(CreateItem.Name);
end.
"#,
    );
    assert_eq!(out, vec!["ItemReturnedFromFunc"]);
}

#[test]
fn test_interfaces_com_supports_guid_string() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ITarget = interface ['{99887766-5544-3322-1100-AABBCCDDEEFF}'] end;
type TTargetImpl = class(TInterfacedObject, ITarget) end;

var obj: TObject;
begin
  obj := TTargetImpl.Create;
  WriteLn(Supports(obj, '{99887766-5544-3322-1100-AABBCCDDEEFF}'));
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_interfaces_corba_nil_check() {
    let out = run_pascal(
        r#"
program Test;
{$INTERFACES CORBA}
type ICorbaNil = interface ['{A1B2C3D4-E5F6-1122-3344-556677889900}'] end;
var c: ICorbaNil;
begin
  c := nil;
  WriteLn(c = nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}
