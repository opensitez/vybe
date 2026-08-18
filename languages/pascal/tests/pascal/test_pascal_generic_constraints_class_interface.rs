use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 85: Generic Constraints (class, record, interface, constructor)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_generic_constraint_class_basic() {
    let out = run_pascal(
        r#"
program Test;
type TClassHolder<T: class> = class
  public Item: T;
  constructor Create(AItem: T);
end;
constructor TClassHolder<T>.Create(AItem: T); begin Item := AItem; end;

type TSampleObj = class end;

var holder: TClassHolder<TSampleObj>; obj: TSampleObj;
begin
  obj := TSampleObj.Create;
  holder := TClassHolder<TSampleObj>.Create(obj);
  WriteLn(holder.Item <> nil);
  holder.Free; obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_generic_constraint_record_basic() {
    let out = run_pascal(
        r#"
program Test;
type TRecordHolder<T: record> = class
  public Item: T;
  constructor Create(const AItem: T);
end;
constructor TRecordHolder<T>.Create(const AItem: T); begin Item := AItem; end;

type TPointRec = record X, Y: Integer; end;

var holder: TRecordHolder<TPointRec>; p: TPointRec;
begin
  p.X := 10; p.Y := 20;
  holder := TRecordHolder<TPointRec>.Create(p);
  WriteLn(holder.Item.X.ToString + ',' + holder.Item.Y.ToString);
  holder.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10,20"]);
}

#[test]
fn test_generic_constraint_constructor() {
    let out = run_pascal(
        r#"
program Test;
type TFactory<T: class, constructor> = class
  public class function CreateInstance: T;
end;
class function TFactory<T>.CreateInstance: T;
begin
  Result := T.Create;
end;

type TWidget = class
  public constructor Create;
end;
constructor TWidget.Create; begin WriteLn('WidgetCreated'); end;

var w: TWidget;
begin
  w := TFactory<TWidget>.CreateInstance;
  w.Free;
end.
"#,
    );
    assert_eq!(out, vec!["WidgetCreated"]);
}

#[test]
fn test_generic_constraint_interface_basic() {
    let out = run_pascal(
        r#"
program Test;
type IWork = interface
  ['{11111111-1111-1111-1111-111111111111}']
  procedure DoWork;
end;

type TWorkInvoker<T: IWork> = class
  public class procedure Run(intf: T);
end;
class procedure TWorkInvoker<T>.Run(intf: T);
begin
  intf.DoWork;
end;

type TWorkImpl = class(TInterfacedObject, IWork)
  public procedure DoWork;
end;
procedure TWorkImpl.DoWork; begin WriteLn('WorkDoneViaConstrainedGeneric'); end;

var w: IWork;
begin
  w := TWorkImpl.Create;
  TWorkInvoker<IWork>.Run(w);
end.
"#,
    );
    assert_eq!(out, vec!["WorkDoneViaConstrainedGeneric"]);
}

#[test]
fn test_generic_constraint_ancestor_class() {
    let out = run_pascal(
        r#"
program Test;
type TBaseObj = class
  public procedure Announce; virtual;
end;
procedure TBaseObj.Announce; begin WriteLn('Base'); end;

type TSubObj = class(TBaseObj)
  public procedure Announce; override;
end;
procedure TSubObj.Announce; begin WriteLn('Sub'); end;

type TRunner<T: TBaseObj> = class
  public class procedure Execute(obj: T);
end;
class procedure TRunner<T>.Execute(obj: T);
begin
  obj.Announce;
end;

var s: TSubObj;
begin
  s := TSubObj.Create;
  TRunner<TSubObj>.Execute(s);
  s.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Sub"]);
}

#[test]
fn test_generic_constraint_multiple_class_interface() {
    let out = run_pascal(
        r#"
program Test;
type ILoggable = interface
  ['{22222222-2222-2222-2222-222222222222}']
  procedure Log;
end;

type TLoggerContainer<T: class, ILoggable> = class
  public class procedure Process(obj: T);
end;
class procedure TLoggerContainer<T>.Process(obj: T);
begin
  obj.Log;
end;

type TLoggableObj = class(TInterfacedObject, ILoggable)
  public procedure Log;
end;
procedure TLoggableObj.Log; begin WriteLn('ObjLogged'); end;

var l: TLoggableObj;
begin
  l := TLoggableObj.Create;
  TLoggerContainer<TLoggableObj>.Process(l);
end.
"#,
    );
    assert_eq!(out, vec!["ObjLogged"]);
}

#[test]
fn test_generic_constraint_record_with_method() {
    let out = run_pascal(
        r#"
program Test;
type TRecHolder<T: record> = class
  public class procedure PrintSize;
end;
class procedure TRecHolder<T>.PrintSize;
begin
  WriteLn(SizeOf(T));
end;

type TMyRec = record A, B: Integer; end;

begin
  TRecHolder<TMyRec>.PrintSize;
end.
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_generic_constraint_interface_method_call() {
    let out = run_pascal(
        r#"
program Test;
type ICalculable = interface
  ['{33333333-3333-3333-3333-333333333333}']
  function GetVal: Integer;
end;

type TCalcRunner<T: ICalculable> = class
  public class function Execute(intf: T): Integer;
end;
class function TCalcRunner<T>.Execute(intf: T): Integer;
begin
  Result := intf.GetVal * 2;
end;

type TCalcImpl = class(TInterfacedObject, ICalculable)
  public function GetVal: Integer;
end;
function TCalcImpl.GetVal: Integer; begin Result := 21; end;

var c: ICalculable;
begin
  c := TCalcImpl.Create;
  WriteLn(TCalcRunner<ICalculable>.Execute(c));
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_generic_constraint_class_is_nil_check() {
    let out = run_pascal(
        r#"
program Test;
type TNullableHolder<T: class> = class
  public class function IsNil(obj: T): Boolean;
end;
class function TNullableHolder<T>.IsNil(obj: T): Boolean;
begin
  Result := obj = nil;
end;

type TDummy = class end;

var d: TDummy;
begin
  d := nil;
  WriteLn(TNullableHolder<TDummy>.IsNil(d));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_generic_constraint_two_type_parameters() {
    let out = run_pascal(
        r#"
program Test;
type TPair<TKey: record; TValue: class> = class
  public Key: TKey; Value: TValue;
  constructor Create(K: TKey; V: TValue);
end;
constructor TPair<TKey, TValue>.Create(K: TKey; V: TValue);
begin
  Key := K; Value := V;
end;

type TValObj = class end;

var p: TPair<Integer, TValObj>; obj: TValObj;
begin
  obj := TValObj.Create;
  p := TPair<Integer, TValObj>.Create(101, obj);
  WriteLn(p.Key);
  p.Free; obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["101"]);
}

#[test]
fn test_generic_constraint_constructor_with_param() {
    let out = run_pascal(
        r#"
program Test;
type TBaseItem = class
  public constructor Create; virtual;
end;
constructor TBaseItem.Create; begin end;

type TDerivedItem = class(TBaseItem)
  public constructor Create; override;
end;
constructor TDerivedItem.Create; begin WriteLn('DerivedCreated'); end;

type TCreator<T: TBaseItem, constructor> = class
  public class function Make: T;
end;
class function TCreator<T>.Make: T;
begin
  Result := T.Create;
end;

var item: TDerivedItem;
begin
  item := TCreator<TDerivedItem>.Make;
  item.Free;
end.
"#,
    );
    assert_eq!(out, vec!["DerivedCreated"]);
}

#[test]
fn test_generic_constraint_interface_multiple_interfaces() {
    let out = run_pascal(
        r#"
program Test;
type IA = interface ['{44444444-4444-4444-4444-444444444444}'] procedure DoA; end;
type IB = interface ['{55555555-5555-5555-5555-555555555555}'] procedure DoB; end;

type TDualHandler<T: IA, IB> = class
  public class procedure Run(intf: T);
end;
class procedure TDualHandler<T>.Run(intf: T);
begin
  intf.DoA;
  intf.DoB;
end;

type TDualImpl = class(TInterfacedObject, IA, IB)
  public procedure DoA; procedure DoB;
end;
procedure TDualImpl.DoA; begin WriteLn('ExecA'); end;
procedure TDualImpl.DoB; begin WriteLn('ExecB'); end;

var impl: TDualImpl;
begin
  impl := TDualImpl.Create;
  TDualHandler<TDualImpl>.Run(impl);
end.
"#,
    );
    assert_eq!(out, vec!["ExecA", "ExecB"]);
}

#[test]
fn test_generic_constraint_class_type_casting() {
    let out = run_pascal(
        r#"
program Test;
type TParent = class end;
type TChild = class(TParent) public procedure SubMethod; end;
procedure TChild.SubMethod; begin WriteLn('ChildMethod'); end;

type TCastHelper<T: TParent> = class
  public class procedure CallIfChild(obj: T);
end;
class procedure TCastHelper<T>.CallIfChild(obj: T);
begin
  if obj is TChild then TChild(obj).SubMethod;
end;

var c: TChild;
begin
  c := TChild.Create;
  TCastHelper<TChild>.CallIfChild(c);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ChildMethod"]);
}

#[test]
fn test_generic_constraint_record_type_info() {
    let out = run_pascal(
        r#"
program Test;
type TRecInspector<T: record> = class
  public class procedure PrintName;
end;
class procedure TRecInspector<T>.PrintName;
begin
  WriteLn(TypeInfo(T)^.Name);
end;

type TSampleRecord = record A: Integer; end;

begin
  TRecInspector<TSampleRecord>.PrintName;
end.
"#,
    );
    assert_eq!(out, vec!["TSampleRecord"]);
}

#[test]
fn test_generic_constraint_interface_property_access() {
    let out = run_pascal(
        r#"
program Test;
type INameable = interface
  ['{66666666-6666-6666-6666-666666666666}']
  function GetName: String;
  property Name: String read GetName;
end;

type TNamePrinter<T: INameable> = class
  public class procedure Print(intf: T);
end;
class procedure TNamePrinter<T>.Print(intf: T);
begin
  WriteLn(intf.Name);
end;

type TNameImpl = class(TInterfacedObject, INameable)
  public function GetName: String;
end;
function TNameImpl.GetName: String; begin Result := 'ConstrainedInterfaceName'; end;

var n: INameable;
begin
  n := TNameImpl.Create;
  TNamePrinter<INameable>.Print(n);
end.
"#,
    );
    assert_eq!(out, vec!["ConstrainedInterfaceName"]);
}

#[test]
fn test_generic_constraint_class_virtual_method_dispatch() {
    let out = run_pascal(
        r#"
program Test;
type TBaseShape = class
  public function Area: Double; virtual;
end;
function TBaseShape.Area: Double; begin Result := 0.0; end;

type TRectangle = class(TBaseShape)
  public W, H: Double;
  constructor Create(AW, AH: Double);
  function Area: Double; override;
end;
constructor TRectangle.Create(AW, AH: Double); begin W := AW; H := AH; end;
function TRectangle.Area: Double; begin Result := W * H; end;

type TAreaCalc<T: TBaseShape> = class
  public class function Compute(shape: T): Double;
end;
class function TAreaCalc<T>.Compute(shape: T): Double;
begin
  Result := shape.Area;
end;

var rect: TRectangle;
begin
  rect := TRectangle.Create(5.0, 4.0);
  WriteLn(TAreaCalc<TRectangle>.Compute(rect));
  rect.Free;
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_generic_constraint_interface_queryinterface() {
    let out = run_pascal(
        r#"
program Test;
type IFirst = interface ['{77777777-7777-7777-7777-777777777777}'] end;
type ISecond = interface ['{88888888-8888-8888-8888-888888888888}'] procedure SecondProc; end;

type TIntfQuery<T: IFirst> = class
  public class procedure CheckSecond(intf: T);
end;
class procedure TIntfQuery<T>.CheckSecond(intf: T);
var sec: ISecond;
begin
  if Supports(intf, ISecond, sec) then sec.SecondProc;
end;

type TMultiImpl = class(TInterfacedObject, IFirst, ISecond)
  public procedure SecondProc;
end;
procedure TMultiImpl.SecondProc; begin WriteLn('SecondSupported'); end;

var m: IFirst;
begin
  m := TMultiImpl.Create;
  TIntfQuery<IFirst>.CheckSecond(m);
end.
"#,
    );
    assert_eq!(out, vec!["SecondSupported"]);
}

#[test]
fn test_generic_constraint_record_in_list() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections;
type TRecStore<T: record> = class
  private FList: TList<T>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(const item: T);
    function Count: Integer;
end;
constructor TRecStore<T>.Create; begin FList := TList<T>.Create; end;
destructor TRecStore<T>.Destroy; begin FList.Free; inherited Destroy; end;
procedure TRecStore<T>.Add(const item: T); begin FList.Add(item); end;
function TRecStore<T>.Count: Integer; begin Result := FList.Count; end;

type TItemRec = record Code: Integer; end;

var store: TRecStore<TItemRec>; r: TItemRec;
begin
  store := TRecStore<TItemRec>.Create;
  r.Code := 1; store.Add(r);
  WriteLn(store.Count);
  store.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_generic_constraint_class_method_generic_parameters() {
    let out = run_pascal(
        r#"
program Test;
type TGenericUtils = class
  public class function CreateAndUse<T: class, constructor>: String;
end;
class function TGenericUtils.CreateAndUse<T>: String;
var obj: T;
begin
  obj := T.Create;
  Result := obj.ClassName;
  obj.Free;
end;

type TTestObj = class end;

begin
  WriteLn(TGenericUtils.CreateAndUse<TTestObj>);
end.
"#,
    );
    assert_eq!(out, vec!["TTestObj"]);
}

#[test]
fn test_generic_constraint_interface_delegation() {
    let out = run_pascal(
        r#"
program Test;
type IRunner = interface ['{99999999-9999-9999-9999-999999999999}'] procedure Run; end;

type TRunnerDelegate<T: IRunner> = class
  private FRunner: T;
  public
    constructor Create(R: T);
    procedure Execute;
end;
constructor TRunnerDelegate<T>.Create(R: T); begin FRunner := R; end;
procedure TRunnerDelegate<T>.Execute; begin FRunner.Run; end;

type TRunnerImpl = class(TInterfacedObject, IRunner)
  public procedure Run;
end;
procedure TRunnerImpl.Run; begin WriteLn('DelegatedRunnerExecuted'); end;

var r: IRunner; delegateObj: TRunnerDelegate<IRunner>;
begin
  r := TRunnerImpl.Create;
  delegateObj := TRunnerDelegate<IRunner>.Create(r);
  delegateObj.Execute;
  delegateObj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["DelegatedRunnerExecuted"]);
}
