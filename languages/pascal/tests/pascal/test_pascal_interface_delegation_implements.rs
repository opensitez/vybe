use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 18: Interface Delegation (implements Clause)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_interface_delegation_basic() {
    let out = run_pascal(
        r#"
program Test;
type ILogger = interface
  ['{11110000-1111-1111-1111-111111111111}']
  procedure Log(msg: String);
end;
type TLoggerImpl = class(TInterfacedObject, ILogger)
  public procedure Log(msg: String);
end;
type TAppController = class(TInterfacedObject, ILogger)
  private FLogger: ILogger;
  public constructor Create;
  public property Logger: ILogger read FLogger implements ILogger;
end;
procedure TLoggerImpl.Log(msg: String); begin WriteLn('DELEGATED: ' + msg); end;
constructor TAppController.Create; begin FLogger := TLoggerImpl.Create; end;
var app: ILogger;
begin
  app := TAppController.Create;
  app.Log('Delegation Works');
end.
"#,
    );
    assert_eq!(out, vec!["DELEGATED: Delegation Works"]);
}

#[test]
fn test_interface_delegation_multiple_methods() {
    let out = run_pascal(
        r#"
program Test;
type IOperations = interface
  ['{22220000-2222-2222-2222-222222222222}']
  function Add(a, b: Integer): Integer;
  function Sub(a, b: Integer): Integer;
end;
type TOpsImpl = class(TInterfacedObject, IOperations)
  public function Add(a, b: Integer): Integer;
  public function Sub(a, b: Integer): Integer;
end;
type TCalculator = class(TInterfacedObject, IOperations)
  private FOps: IOperations;
  public constructor Create;
  public property Ops: IOperations read FOps implements IOperations;
end;
function TOpsImpl.Add(a, b: Integer): Integer; begin Result := a + b; end;
function TOpsImpl.Sub(a, b: Integer): Integer; begin Result := a - b; end;
constructor TCalculator.Create; begin FOps := TOpsImpl.Create; end;
var calc: IOperations;
begin
  calc := TCalculator.Create;
  WriteLn(calc.Add(10, 5));
  WriteLn(calc.Sub(10, 5));
end.
"#,
    );
    assert_eq!(out, vec!["15", "5"]);
}

#[test]
fn test_interface_delegation_dynamic_replacement() {
    let out = run_pascal(
        r#"
program Test;
type IPrinter = interface
  ['{33330000-3333-3333-3333-333333333333}']
  procedure Print;
end;
type TPrinterA = class(TInterfacedObject, IPrinter)
  public procedure Print;
end;
type TPrinterB = class(TInterfacedObject, IPrinter)
  public procedure Print;
end;
type TSwitchable = class(TInterfacedObject, IPrinter)
  private FPrinter: IPrinter;
  public constructor Create; procedure SetPrinter(p: IPrinter);
  public property CurrentPrinter: IPrinter read FPrinter implements IPrinter;
end;
procedure TPrinterA.Print; begin WriteLn('PrinterA'); end;
procedure TPrinterB.Print; begin WriteLn('PrinterB'); end;
constructor TSwitchable.Create; begin FPrinter := TPrinterA.Create; end;
procedure TSwitchable.SetPrinter(p: IPrinter); begin FPrinter := p; end;
var s: TSwitchable; intf: IPrinter;
begin
  s := TSwitchable.Create;
  intf := s;
  intf.Print;
  s.SetPrinter(TPrinterB.Create);
  intf.Print;
end.
"#,
    );
    assert_eq!(out, vec!["PrinterA", "PrinterB"]);
}

#[test]
fn test_interface_delegation_getter_backing() {
    let out = run_pascal(
        r#"
program Test;
type IWorker = interface
  ['{44440000-4444-4444-4444-444444444444}']
  procedure Work;
end;
type TWorkerImpl = class(TInterfacedObject, IWorker)
  public procedure Work;
end;
type TManager = class(TInterfacedObject, IWorker)
  private FWorker: IWorker;
  private function GetWorker: IWorker;
  public constructor Create;
  public property Worker: IWorker read GetWorker implements IWorker;
end;
procedure TWorkerImpl.Work; begin WriteLn('WorkingLazy'); end;
constructor TManager.Create; begin FWorker := nil; end;
function TManager.GetWorker: IWorker;
begin
  if FWorker = nil then FWorker := TWorkerImpl.Create;
  Result := FWorker;
end;
var m: IWorker;
begin
  m := TManager.Create;
  m.Work;
end.
"#,
    );
    assert_eq!(out, vec!["WorkingLazy"]);
}

#[test]
fn test_interface_delegation_multiple_interfaces() {
    let out = run_pascal(
        r#"
program Test;
type INameable = interface
  ['{55550000-5555-5555-5555-555555555555}']
  function GetName: String;
end;
type IValuable = interface
  ['{66660000-6666-6666-6666-666666666666}']
  function GetValue: Integer;
end;
type TNameImpl = class(TInterfacedObject, INameable)
  public function GetName: String;
end;
type TValImpl = class(TInterfacedObject, IValuable)
  public function GetValue: Integer;
end;
type TComposite = class(TInterfacedObject, INameable, IValuable)
  private FName: INameable; FVal: IValuable;
  public constructor Create;
  public property NameImpl: INameable read FName implements INameable;
  public property ValImpl: IValuable read FVal implements IValuable;
end;
function TNameImpl.GetName: String; begin Result := 'CompositeItem'; end;
function TValImpl.GetValue: Integer; begin Result := 999; end;
constructor TComposite.Create; begin FName := TNameImpl.Create; FVal := TValImpl.Create; end;
var comp: TComposite; n: INameable; v: IValuable;
begin
  comp := TComposite.Create;
  n := comp as INameable;
  v := comp as IValuable;
  WriteLn(n.GetName);
  WriteLn(v.GetValue);
end.
"#,
    );
    assert_eq!(out, vec!["CompositeItem", "999"]);
}

#[test]
fn test_interface_delegation_supports_check() {
    let out = run_pascal(
        r#"
program Test;
type IAuditable = interface
  ['{77770000-7777-7777-7777-777777777777}']
  procedure Audit;
end;
type TAuditImpl = class(TInterfacedObject, IAuditable)
  public procedure Audit;
end;
type TDelegatedAudit = class(TInterfacedObject, IAuditable)
  private FAudit: IAuditable;
  public constructor Create;
  public property AuditImpl: IAuditable read FAudit implements IAuditable;
end;
procedure TAuditImpl.Audit; begin WriteLn('AuditExecuted'); end;
constructor TDelegatedAudit.Create; begin FAudit := TAuditImpl.Create; end;
var obj: TObject; intf: IAuditable;
begin
  obj := TDelegatedAudit.Create;
  if Supports(obj, IAuditable, intf) then
    intf.Audit;
end.
"#,
    );
    assert_eq!(out, vec!["AuditExecuted"]);
}

#[test]
fn test_interface_delegation_with_var_parameter() {
    let out = run_pascal(
        r#"
program Test;
type IMutator = interface
  ['{88880000-8888-8888-8888-888888888888}']
  procedure Modify(var x: Integer);
end;
type TMutatorImpl = class(TInterfacedObject, IMutator)
  public procedure Modify(var x: Integer);
end;
type TMutatorDelegate = class(TInterfacedObject, IMutator)
  private FImpl: IMutator;
  public constructor Create;
  public property Impl: IMutator read FImpl implements IMutator;
end;
procedure TMutatorImpl.Modify(var x: Integer); begin x := x * 3; end;
constructor TMutatorDelegate.Create; begin FImpl := TMutatorImpl.Create; end;
var m: IMutator; val: Integer;
begin
  m := TMutatorDelegate.Create;
  val := 10;
  m.Modify(val);
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_interface_delegation_with_record_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TRec = record Val: String; end;
type IRecHandler = interface
  ['{99990000-9999-9999-9999-999999999999}']
  procedure HandleRec(r: TRec);
end;
type TRecHandlerImpl = class(TInterfacedObject, IRecHandler)
  public procedure HandleRec(r: TRec);
end;
type TRecDelegate = class(TInterfacedObject, IRecHandler)
  private FImpl: IRecHandler;
  public constructor Create;
  public property Impl: IRecHandler read FImpl implements IRecHandler;
end;
procedure TRecHandlerImpl.HandleRec(r: TRec); begin WriteLn('REC:' + r.Val); end;
constructor TRecDelegate.Create; begin FImpl := TRecHandlerImpl.Create; end;
var h: IRecHandler; data: TRec;
begin
  data.Val := 'DelegatedRecordData';
  h := TRecDelegate.Create;
  h.HandleRec(data);
end.
"#,
    );
    assert_eq!(out, vec!["REC:DelegatedRecordData"]);
}

#[test]
fn test_interface_delegation_with_default_parameter() {
    let out = run_pascal(
        r#"
program Test;
type IDefaultProc = interface
  ['{AAAA0000-AAAA-AAAA-AAAA-AAAAAAAAAAAA}']
  procedure Run(prefix: String = 'DEF');
end;
type TDefaultImpl = class(TInterfacedObject, IDefaultProc)
  public procedure Run(prefix: String);
end;
type TDefaultDelegate = class(TInterfacedObject, IDefaultProc)
  private FImpl: IDefaultProc;
  public constructor Create;
  public property Impl: IDefaultProc read FImpl implements IDefaultProc;
end;
procedure TDefaultImpl.Run(prefix: String); begin WriteLn('RUN:' + prefix); end;
constructor TDefaultDelegate.Create; begin FImpl := TDefaultImpl.Create; end;
var d: IDefaultProc;
begin
  d := TDefaultDelegate.Create;
  d.Run;
  d.Run('CUST');
end.
"#,
    );
    assert_eq!(out, vec!["RUN:DEF", "RUN:CUST"]);
}

#[test]
fn test_interface_delegation_returning_boolean() {
    let out = run_pascal(
        r#"
program Test;
type ICheck = interface
  ['{BBBB0000-BBBB-BBBB-BBBB-BBBBBBBBBBBB}']
  function Check(val: Integer): Boolean;
end;
type TCheckImpl = class(TInterfacedObject, ICheck)
  public function Check(val: Integer): Boolean;
end;
type TCheckDelegate = class(TInterfacedObject, ICheck)
  private FImpl: ICheck;
  public constructor Create;
  public property Impl: ICheck read FImpl implements ICheck;
end;
function TCheckImpl.Check(val: Integer): Boolean; begin Result := val > 50; end;
constructor TCheckDelegate.Create; begin FImpl := TCheckImpl.Create; end;
var c: ICheck;
begin
  c := TCheckDelegate.Create;
  WriteLn(c.Check(75));
  WriteLn(c.Check(25));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_interface_delegation_modifies_delegate_state() {
    let out = run_pascal(
        r#"
program Test;
type ICounter = interface
  ['{CCCC0000-CCCC-CCCC-CCCC-CCCCCCCCCCCC}']
  procedure IncCount;
  function GetCount: Integer;
end;
type TCounterImpl = class(TInterfacedObject, ICounter)
  private FCount: Integer;
  public procedure IncCount; function GetCount: Integer;
end;
type TCounterWrapper = class(TInterfacedObject, ICounter)
  private FImpl: ICounter;
  public constructor Create;
  public property Impl: ICounter read FImpl implements ICounter;
end;
procedure TCounterImpl.IncCount; begin Inc(FCount); end;
function TCounterImpl.GetCount: Integer; begin Result := FCount; end;
constructor TCounterWrapper.Create; begin FImpl := TCounterImpl.Create; end;
var c: ICounter;
begin
  c := TCounterWrapper.Create;
  c.IncCount;
  c.IncCount;
  WriteLn(c.GetCount);
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_interface_delegation_subclass_extension() {
    let out = run_pascal(
        r#"
program Test;
type IService = interface
  ['{DDDD0000-DDDD-DDDD-DDDD-DDDDDDDDDDDD}']
  procedure Execute;
end;
type TServiceImpl = class(TInterfacedObject, IService)
  public procedure Execute;
end;
type TBaseContainer = class(TInterfacedObject, IService)
  private FService: IService;
  public constructor Create;
  public property Service: IService read FService implements IService;
end;
type TDerivedContainer = class(TBaseContainer) end;
procedure TServiceImpl.Execute; begin WriteLn('SubclassDelegated'); end;
constructor TBaseContainer.Create; begin FService := TServiceImpl.Create; end;
var s: IService;
begin
  s := TDerivedContainer.Create;
  s.Execute;
end.
"#,
    );
    assert_eq!(out, vec!["SubclassDelegated"]);
}

#[test]
fn test_interface_delegation_overloaded_methods() {
    let out = run_pascal(
        r#"
program Test;
type IOverloaded = interface
  ['{EEEE0000-EEEE-EEEE-EEEE-EEEEEEEEEEEE}']
  procedure Process(i: Integer); overload;
  procedure Process(s: String); overload;
end;
type TOverloadedImpl = class(TInterfacedObject, IOverloaded)
  public procedure Process(i: Integer); overload;
  public procedure Process(s: String); overload;
end;
type TOverloadedDelegate = class(TInterfacedObject, IOverloaded)
  private FImpl: IOverloaded;
  public constructor Create;
  public property Impl: IOverloaded read FImpl implements IOverloaded;
end;
procedure TOverloadedImpl.Process(i: Integer); begin WriteLn('INT:' + i.ToString); end;
procedure TOverloadedImpl.Process(s: String); begin WriteLn('STR:' + s); end;
constructor TOverloadedDelegate.Create; begin FImpl := TOverloadedImpl.Create; end;
var o: IOverloaded;
begin
  o := TOverloadedDelegate.Create;
  o.Process(42);
  o.Process('Pascal');
end.
"#,
    );
    assert_eq!(out, vec!["INT:42", "STR:Pascal"]);
}

#[test]
fn test_interface_delegation_enum_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TMode = (mOne, mTwo);
type IModeHandler = interface
  ['{FFFF0000-FFFF-FFFF-FFFF-FFFFFFFFFFFF}']
  procedure SetMode(m: TMode);
end;
type TModeImpl = class(TInterfacedObject, IModeHandler)
  public procedure SetMode(m: TMode);
end;
type TModeDelegate = class(TInterfacedObject, IModeHandler)
  private FImpl: IModeHandler;
  public constructor Create;
  public property Impl: IModeHandler read FImpl implements IModeHandler;
end;
procedure TModeImpl.SetMode(m: TMode); begin WriteLn('MODE:' + Ord(m).ToString); end;
constructor TModeDelegate.Create; begin FImpl := TModeImpl.Create; end;
var mh: IModeHandler;
begin
  mh := TModeDelegate.Create;
  mh.SetMode(mTwo);
end.
"#,
    );
    assert_eq!(out, vec!["MODE:1"]);
}

#[test]
fn test_interface_delegation_nested_composition() {
    let out = run_pascal(
        r#"
program Test;
type IAction = interface
  ['{10100000-1010-1010-1010-101010101010}']
  procedure Run;
end;
type TCoreAction = class(TInterfacedObject, IAction)
  public procedure Run;
end;
type TMiddleWrapper = class(TInterfacedObject, IAction)
  private FCore: IAction;
  public constructor Create;
  public property Core: IAction read FCore implements IAction;
end;
type TOuterWrapper = class(TInterfacedObject, IAction)
  private FMid: IAction;
  public constructor Create;
  public property Mid: IAction read FMid implements IAction;
end;
procedure TCoreAction.Run; begin WriteLn('DeepNestedRun'); end;
constructor TMiddleWrapper.Create; begin FCore := TCoreAction.Create; end;
constructor TOuterWrapper.Create; begin FMid := TMiddleWrapper.Create; end;
var a: IAction;
begin
  a := TOuterWrapper.Create;
  a.Run;
end.
"#,
    );
    assert_eq!(out, vec!["DeepNestedRun"]);
}

#[test]
fn test_interface_delegation_returning_real() {
    let out = run_pascal(
        r#"
program Test;
type IMath = interface
  ['{20200000-2020-2020-2020-202020202020}']
  function SqrtVal(x: Real): Real;
end;
type TMathImpl = class(TInterfacedObject, IMath)
  public function SqrtVal(x: Real): Real;
end;
type TMathDelegate = class(TInterfacedObject, IMath)
  private FImpl: IMath;
  public constructor Create;
  public property Impl: IMath read FImpl implements IMath;
end;
function TMathImpl.SqrtVal(x: Real): Real; begin Result := Sqrt(x); end;
constructor TMathDelegate.Create; begin FImpl := TMathImpl.Create; end;
var m: IMath;
begin
  m := TMathDelegate.Create;
  WriteLn(m.SqrtVal(16.0));
end.
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_interface_delegation_const_parameter() {
    let out = run_pascal(
        r#"
program Test;
type IDisplay = interface
  ['{30300000-3030-3030-3030-303030303030}']
  procedure ShowConst(const s: String);
end;
type TDisplayImpl = class(TInterfacedObject, IDisplay)
  public procedure ShowConst(const s: String);
end;
type TDisplayDelegate = class(TInterfacedObject, IDisplay)
  private FImpl: IDisplay;
  public constructor Create;
  public property Impl: IDisplay read FImpl implements IDisplay;
end;
procedure TDisplayImpl.ShowConst(const s: String); begin WriteLn('CONST:' + s); end;
constructor TDisplayDelegate.Create; begin FImpl := TDisplayImpl.Create; end;
var d: IDisplay;
begin
  d := TDisplayDelegate.Create;
  d.ShowConst('ImmutableText');
end.
"#,
    );
    assert_eq!(out, vec!["CONST:ImmutableText"]);
}

#[test]
fn test_interface_delegation_array_returning_method() {
    let out = run_pascal(
        r#"
program Test;
type TArr = array[1..3] of Integer;
type IArrayProvider = interface
  ['{40400000-4040-4040-4040-404040404040}']
  function GetArray: TArr;
end;
type TArrayImpl = class(TInterfacedObject, IArrayProvider)
  public function GetArray: TArr;
end;
type TArrayDelegate = class(TInterfacedObject, IArrayProvider)
  private FImpl: IArrayProvider;
  public constructor Create;
  public property Impl: IArrayProvider read FImpl implements IArrayProvider;
end;
function TArrayImpl.GetArray: TArr; begin Result[1] := 10; Result[2] := 20; Result[3] := 30; end;
constructor TArrayDelegate.Create; begin FImpl := TArrayImpl.Create; end;
var ap: IArrayProvider; arr: TArr;
begin
  ap := TArrayDelegate.Create;
  arr := ap.GetArray;
  WriteLn(arr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_interface_delegation_subrange_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 1..100;
type IScoreHandler = interface
  ['{50500000-5050-5050-5050-505050505050}']
  procedure RecordScore(s: TScore);
end;
type TScoreImpl = class(TInterfacedObject, IScoreHandler)
  public procedure RecordScore(s: TScore);
end;
type TScoreDelegate = class(TInterfacedObject, IScoreHandler)
  private FImpl: IScoreHandler;
  public constructor Create;
  public property Impl: IScoreHandler read FImpl implements IScoreHandler;
end;
procedure TScoreImpl.RecordScore(s: TScore); begin WriteLn('SCORE:' + s.ToString); end;
constructor TScoreDelegate.Create; begin FImpl := TScoreImpl.Create; end;
var sh: IScoreHandler; sc: TScore;
begin
  sc := 95;
  sh := TScoreDelegate.Create;
  sh.RecordScore(sc);
end.
"#,
    );
    assert_eq!(out, vec!["SCORE:95"]);
}

#[test]
fn test_interface_delegation_ref_count_tracking() {
    let out = run_pascal(
        r#"
program Test;
type ICounterIntf = interface
  ['{60600000-6060-6060-6060-606060606060}']
  procedure Ping;
end;
type TCounterImpl = class(TInterfacedObject, ICounterIntf)
  public procedure Ping; destructor Destroy; override;
end;
type TCounterDelegate = class(TInterfacedObject, ICounterIntf)
  private FImpl: ICounterIntf;
  public constructor Create;
  public property Impl: ICounterIntf read FImpl implements ICounterIntf;
end;
procedure TCounterImpl.Ping; begin WriteLn('DelegatedPing'); end;
destructor TCounterImpl.Destroy; begin WriteLn('DelegateDestroyed'); inherited Destroy; end;
constructor TCounterDelegate.Create; begin FImpl := TCounterImpl.Create; end;
procedure TestLifetime;
var d: ICounterIntf;
begin
  d := TCounterDelegate.Create;
  d.Ping;
end;
begin
  TestLifetime;
end.
"#,
    );
    assert_eq!(out, vec!["DelegatedPing", "DelegateDestroyed"]);
}
