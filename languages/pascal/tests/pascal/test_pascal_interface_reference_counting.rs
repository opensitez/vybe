use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 17: Interface Implementations & Reference Counting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_interface_single_implementation() {
    let out = run_pascal(r#"
program Test;
type ILogger = interface
  ['{11111111-1111-1111-1111-111111111111}']
  procedure Log(msg: String);
end;
type TConsoleLogger = class(TInterfacedObject, ILogger)
  public procedure Log(msg: String);
end;
procedure TConsoleLogger.Log(msg: String);
begin
  WriteLn('LOG: ' + msg);
end;
var logger: ILogger;
begin
  logger := TConsoleLogger.Create;
  logger.Log('Interface System Init');
end.
"#);
    assert_eq!(out, vec!["LOG: Interface System Init"]);
}

#[test]
fn test_interface_auto_reference_counting_cleanup() {
    let out = run_pascal(r#"
program Test;
type ITracker = interface
  ['{22222222-2222-2222-2222-222222222222}']
  procedure Ping;
end;
type TTrackerObj = class(TInterfacedObject, ITracker)
  public procedure Ping;
  public destructor Destroy; override;
end;
procedure TTrackerObj.Ping; begin WriteLn('Pinged'); end;
destructor TTrackerObj.Destroy; begin WriteLn('AutoDestroyed'); inherited Destroy; end;
procedure RunScope;
var t: ITracker;
begin
  t := TTrackerObj.Create;
  t.Ping;
end;
begin
  RunScope;
end.
"#);
    assert_eq!(out, vec!["Pinged", "AutoDestroyed"]);
}

#[test]
fn test_multiple_interfaces_on_single_class() {
    let out = run_pascal(r#"
program Test;
type IReader = interface
  ['{33333333-3333-3333-3333-333333333333}']
  function ReadData: String;
end;
type IWriter = interface
  ['{44444444-4444-4444-4444-444444444444}']
  procedure WriteData(s: String);
end;
type TFileHandler = class(TInterfacedObject, IReader, IWriter)
  private FBuffer: String;
  public function ReadData: String;
  public procedure WriteData(s: String);
end;
function TFileHandler.ReadData: String; begin Result := FBuffer; end;
procedure TFileHandler.WriteData(s: String); begin FBuffer := s; end;
var r: IReader; w: IWriter; h: TFileHandler;
begin
  h := TFileHandler.Create;
  w := h; r := h;
  w.WriteData('StreamContent');
  WriteLn(r.ReadData);
end.
"#);
    assert_eq!(out, vec!["StreamContent"]);
}

#[test]
fn test_interface_parameter_passing() {
    let out = run_pascal(r#"
program Test;
type IPrintable = interface
  ['{55555555-5555-5555-5555-555555555555}']
  function GetText: String;
end;
type TReport = class(TInterfacedObject, IPrintable)
  public function GetText: String;
end;
function TReport.GetText: String; begin Result := 'ReportContent'; end;
procedure PrintObject(p: IPrintable);
begin
  WriteLn('OUTPUT: ' + p.GetText);
end;
var rep: IPrintable;
begin
  rep := TReport.Create;
  PrintObject(rep);
end.
"#);
    assert_eq!(out, vec!["OUTPUT: ReportContent"]);
}

#[test]
fn test_interface_function_return_value() {
    let out = run_pascal(r#"
program Test;
type IWorker = interface
  ['{66666666-6666-6666-6666-666666666666}']
  procedure DoWork;
end;
type TWorkerImpl = class(TInterfacedObject, IWorker)
  public procedure DoWork;
end;
procedure TWorkerImpl.DoWork; begin WriteLn('WorkDone'); end;
function CreateWorker: IWorker;
begin
  Result := TWorkerImpl.Create;
end;
var w: IWorker;
begin
  w := CreateWorker;
  w.DoWork;
end.
"#);
    assert_eq!(out, vec!["WorkDone"]);
}

#[test]
fn test_interface_array_polymorphism() {
    let out = run_pascal(r#"
program Test;
type ITask = interface
  ['{77777777-7777-7777-7777-777777777777}']
  procedure Execute;
end;
type TTaskA = class(TInterfacedObject, ITask)
  public procedure Execute;
end;
type TTaskB = class(TInterfacedObject, ITask)
  public procedure Execute;
end;
procedure TTaskA.Execute; begin WriteLn('TaskA'); end;
procedure TTaskB.Execute; begin WriteLn('TaskB'); end;
var tasks: array[1..2] of ITask; i: Integer;
begin
  tasks[1] := TTaskA.Create;
  tasks[2] := TTaskB.Create;
  for i := 1 to 2 do
    tasks[i].Execute;
end.
"#);
    assert_eq!(out, vec!["TaskA", "TaskB"]);
}

#[test]
fn test_supports_operator_interface_query() {
    let out = run_pascal(r#"
program Test;
type IAuditable = interface
  ['{88888888-8888-8888-8888-888888888888}']
  procedure Audit;
end;
type TAuditService = class(TInterfacedObject, IAuditable)
  public procedure Audit;
end;
procedure TAuditService.Audit; begin WriteLn('Audited'); end;
var obj: TObject; aud: IAuditable;
begin
  obj := TAuditService.Create;
  if Supports(obj, IAuditable, aud) then
    aud.Audit;
end.
"#);
    assert_eq!(out, vec!["Audited"]);
}

#[test]
fn test_interface_inheritance() {
    let out = run_pascal(r#"
program Test;
type IBaseIntf = interface
  ['{99999999-9999-9999-9999-999999999999}']
  procedure Step1;
end;
type ISubIntf = interface(IBaseIntf)
  ['{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}']
  procedure Step2;
end;
type TChainImpl = class(TInterfacedObject, ISubIntf)
  public procedure Step1; procedure Step2;
end;
procedure TChainImpl.Step1; begin WriteLn('Step1'); end;
procedure TChainImpl.Step2; begin WriteLn('Step2'); end;
var sub: ISubIntf;
begin
  sub := TChainImpl.Create;
  sub.Step1;
  sub.Step2;
end.
"#);
    assert_eq!(out, vec!["Step1", "Step2"]);
}

#[test]
fn test_interface_reassigned_to_nil() {
    let out = run_pascal(r#"
program Test;
type ICleanable = interface
  ['{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}']
  procedure Action;
end;
type TCleanObj = class(TInterfacedObject, ICleanable)
  public procedure Action; destructor Destroy; override;
end;
procedure TCleanObj.Action; begin WriteLn('Active'); end;
destructor TCleanObj.Destroy; begin WriteLn('DestroyedOnNil'); inherited Destroy; end;
var intf: ICleanable;
begin
  intf := TCleanObj.Create;
  intf.Action;
  intf := nil;
  WriteLn('AfterNil');
end.
"#);
    assert_eq!(out, vec!["Active", "DestroyedOnNil", "AfterNil"]);
}

#[test]
fn test_interface_property_accessors() {
    let out = run_pascal(r#"
program Test;
type IHasName = interface
  ['{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}']
  function GetName: String;
  procedure SetName(s: String);
  property Name: String read GetName write SetName;
end;
type TNamedObj = class(TInterfacedObject, IHasName)
  private FName: String;
  public function GetName: String; procedure SetName(s: String);
end;
function TNamedObj.GetName: String; begin Result := FName; end;
procedure TNamedObj.SetName(s: String); begin FName := s; end;
var item: IHasName;
begin
  item := TNamedObj.Create;
  item.Name := 'InterfaceName';
  WriteLn(item.Name);
end.
"#);
    assert_eq!(out, vec!["InterfaceName"]);
}

#[test]
fn test_interface_ref_count_shared_ownership() {
    let out = run_pascal(r#"
program Test;
type IShared = interface
  ['{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}']
  procedure Show;
end;
type TSharedObj = class(TInterfacedObject, IShared)
  public procedure Show; destructor Destroy; override;
end;
procedure TSharedObj.Show; begin WriteLn('SharedShow'); end;
destructor TSharedObj.Destroy; begin WriteLn('SharedDestroyed'); inherited Destroy; end;
var ref1, ref2: IShared;
begin
  ref1 := TSharedObj.Create;
  ref2 := ref1;
  ref1 := nil;
  WriteLn('Ref1Cleared');
  ref2.Show;
  ref2 := nil;
  WriteLn('Ref2Cleared');
end.
"#);
    assert_eq!(out, vec!["Ref1Cleared", "SharedShow", "SharedDestroyed", "Ref2Cleared"]);
}

#[test]
fn test_interface_method_returning_integer() {
    let out = run_pascal(r#"
program Test;
type ICalculator = interface
  ['{EEEEEEEE-EEEE-EEEE-EEEE-EEEEEEEEEEEE}']
  function Multiply(a, b: Integer): Integer;
end;
type TSimpleCalc = class(TInterfacedObject, ICalculator)
  public function Multiply(a, b: Integer): Integer;
end;
function TSimpleCalc.Multiply(a, b: Integer): Integer; begin Result := a * b; end;
var calc: ICalculator;
begin
  calc := TSimpleCalc.Create;
  WriteLn(calc.Multiply(6, 7));
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_interface_method_returning_boolean() {
    let out = run_pascal(r#"
program Test;
type IChecker = interface
  ['{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}']
  function IsValid(s: String): Boolean;
end;
type TLengthChecker = class(TInterfacedObject, IChecker)
  public function IsValid(s: String): Boolean;
end;
function TLengthChecker.IsValid(s: String): Boolean; begin Result := Length(s) >= 3; end;
var c: IChecker;
begin
  c := TLengthChecker.Create;
  WriteLn(c.IsValid('Pascal'));
  WriteLn(c.IsValid('Hi'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_interface_assigned_check() {
    let out = run_pascal(r#"
program Test;
type IEvent = interface
  ['{10101010-1010-1010-1010-101010101010}']
  procedure Trigger;
end;
var e: IEvent;
begin
  WriteLn(Assigned(e));
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_interface_with_var_parameter() {
    let out = run_pascal(r#"
program Test;
type ITransformer = interface
  ['{20202020-2020-2020-2020-202020202020}']
  procedure DoubleVal(var v: Integer);
end;
type TDoubleImpl = class(TInterfacedObject, ITransformer)
  public procedure DoubleVal(var v: Integer);
end;
procedure TDoubleImpl.DoubleVal(var v: Integer); begin v := v * 2; end;
var t: ITransformer; val: Integer;
begin
  t := TDoubleImpl.Create;
  val := 15;
  t.DoubleVal(val);
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_interface_with_enum_parameter() {
    let out = run_pascal(r#"
program Test;
type TDirection = (dUp, dDown);
type IMover = interface
  ['{30303030-3030-3030-3030-303030303030}']
  procedure Move(dir: TDirection);
end;
type TRobot = class(TInterfacedObject, IMover)
  public procedure Move(dir: TDirection);
end;
procedure TRobot.Move(dir: TDirection); begin WriteLn('Dir:' + Ord(dir).ToString); end;
var m: IMover;
begin
  m := TRobot.Create;
  m.Move(dDown);
end.
"#);
    assert_eq!(out, vec!["Dir:1"]);
}

#[test]
fn test_interface_with_record_parameter() {
    let out = run_pascal(r#"
program Test;
type TData = record Code: Integer; end;
type IDataConsumer = interface
  ['{40404040-4040-4040-4040-404040404040}']
  procedure Consume(d: TData);
end;
type TConsumerImpl = class(TInterfacedObject, IDataConsumer)
  public procedure Consume(d: TData);
end;
procedure TConsumerImpl.Consume(d: TData); begin WriteLn(d.Code); end;
var c: IDataConsumer; rec: TData;
begin
  rec.Code := 500;
  c := TConsumerImpl.Create;
  c.Consume(rec);
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_interface_with_default_parameter() {
    let out = run_pascal(r#"
program Test;
type INotifier = interface
  ['{50505050-5050-5050-5050-505050505050}']
  procedure Notify(msg: String = 'DefaultNotification');
end;
type TNotifierImpl = class(TInterfacedObject, INotifier)
  public procedure Notify(msg: String);
end;
procedure TNotifierImpl.Notify(msg: String); begin WriteLn(msg); end;
var n: INotifier;
begin
  n := TNotifierImpl.Create;
  n.Notify;
  n.Notify('CustomNotification');
end.
"#);
    assert_eq!(out, vec!["DefaultNotification", "CustomNotification"]);
}

#[test]
fn test_interface_cast_from_class_instance() {
    let out = run_pascal(r#"
program Test;
type IService = interface
  ['{60606060-6060-6060-6060-606060606060}']
  procedure Serve;
end;
type TServiceImpl = class(TInterfacedObject, IService)
  public procedure Serve;
end;
procedure TServiceImpl.Serve; begin WriteLn('Served'); end;
var obj: TServiceImpl; srv: IService;
begin
  obj := TServiceImpl.Create;
  srv := obj as IService;
  srv.Serve;
end.
"#);
    assert_eq!(out, vec!["Served"]);
}

#[test]
fn test_interface_factory_registry_dispatch() {
    let out = run_pascal(r#"
program Test;
type ICommand = interface
  ['{70707070-7070-7070-7070-707070707070}']
  procedure Execute;
end;
type TStartCommand = class(TInterfacedObject, ICommand)
  public procedure Execute;
end;
procedure TStartCommand.Execute; begin WriteLn('StartedCommand'); end;
function GetCommand(name: String): ICommand;
begin
  if name = 'start' then Result := TStartCommand.Create else Result := nil;
end;
var cmd: ICommand;
begin
  cmd := GetCommand('start');
  if cmd <> nil then cmd.Execute;
end.
"#);
    assert_eq!(out, vec!["StartedCommand"]);
}
