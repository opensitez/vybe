use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 59: Safecall Calling Convention & Exception Marshaling
// ═══════════════════════════════════════════════════════════

#[test]
fn test_safecall_interface_method_basic() {
    let out = run_pascal(r#"
program Test;
type ISafeIntf = interface
  ['{11111111-2222-3333-4444-555555555555}']
  procedure DoSafeWork; safecall;
end;
type TSafeImpl = class(TInterfacedObject, ISafeIntf)
  public procedure DoSafeWork; safecall;
end;
procedure TSafeImpl.DoSafeWork;
begin
  WriteLn('SafecallExecuted');
end;
var s: ISafeIntf;
begin
  s := TSafeImpl.Create;
  s.DoSafeWork;
end.
"#);
    assert_eq!(out, vec!["SafecallExecuted"]);
}

#[test]
fn test_safecall_exception_marshaling_to_caller() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ISafeErrorIntf = interface
  ['{22222222-3333-4444-5555-666666666666}']
  procedure FailWork; safecall;
end;
type TSafeErrorImpl = class(TInterfacedObject, ISafeErrorIntf)
  public procedure FailWork; safecall;
end;
procedure TSafeErrorImpl.FailWork;
begin
  raise Exception.Create('SafecallErrorOccurred');
end;
var s: ISafeErrorIntf;
begin
  s := TSafeErrorImpl.Create;
  try
    s.FailWork;
  except
    on E: Exception do WriteLn('CallerCaughtSafecall:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["CallerCaughtSafecall:SafecallErrorOccurred"]);
}

#[test]
fn test_safecall_function_return_value() {
    let out = run_pascal(r#"
program Test;
type ISafeCalc = interface
  ['{33333333-4444-5555-6666-777777777777}']
  function Add(a, b: Integer): Integer; safecall;
end;
type TSafeCalcImpl = class(TInterfacedObject, ISafeCalc)
  public function Add(a, b: Integer): Integer; safecall;
end;
function TSafeCalcImpl.Add(a, b: Integer): Integer;
begin
  Result := a + b;
end;
var calc: ISafeCalc;
begin
  calc := TSafeCalcImpl.Create;
  WriteLn(calc.Add(20, 30));
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_safecall_out_parameter() {
    let out = run_pascal(r#"
program Test;
type ISafeOut = interface
  ['{44444444-5555-6666-7777-888888888888}']
  procedure GetInfo(out val: String); safecall;
end;
type TSafeOutImpl = class(TInterfacedObject, ISafeOut)
  public procedure GetInfo(out val: String); safecall;
end;
procedure TSafeOutImpl.GetInfo(out val: String);
begin
  val := 'SafecallOutText';
end;
var obj: ISafeOut; text: String;
begin
  obj := TSafeOutImpl.Create;
  obj.GetInfo(text);
  WriteLn(text);
end.
"#);
    assert_eq!(out, vec!["SafecallOutText"]);
}

#[test]
fn test_safecall_var_parameter() {
    let out = run_pascal(r#"
program Test;
type ISafeMutator = interface
  ['{55555555-6666-7777-8888-999999999999}']
  procedure MultiplyVar(var num: Integer); safecall;
end;
type TSafeMutatorImpl = class(TInterfacedObject, ISafeMutator)
  public procedure MultiplyVar(var num: Integer); safecall;
end;
procedure TSafeMutatorImpl.MultiplyVar(var num: Integer);
begin
  num := num * 5;
end;
var obj: ISafeMutator; val: Integer;
begin
  obj := TSafeMutatorImpl.Create;
  val := 10;
  obj.MultiplyVar(val);
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_safecall_record_parameter() {
    let out = run_pascal(r#"
program Test;
type TDataRec = record Code: Integer; Msg: String; end;
type ISafeRecHandler = interface
  ['{66666666-7777-8888-9999-000000000000}']
  procedure ProcessRec(r: TDataRec); safecall;
end;
type TSafeRecImpl = class(TInterfacedObject, ISafeRecHandler)
  public procedure ProcessRec(r: TDataRec); safecall;
end;
procedure TSafeRecImpl.ProcessRec(r: TDataRec);
begin
  WriteLn(r.Code.ToString + ':' + r.Msg);
end;
var obj: ISafeRecHandler; rec: TDataRec;
begin
  rec.Code := 100; rec.Msg := 'RecData';
  obj := TSafeRecImpl.Create;
  obj.ProcessRec(rec);
end.
"#);
    assert_eq!(out, vec!["100:RecData"]);
}

#[test]
fn test_safecall_enum_parameter() {
    let out = run_pascal(r#"
program Test;
type TMode = (mOff, mOn);
type ISafeMode = interface
  ['{77777777-8888-9999-0000-111111111111}']
  procedure SetMode(m: TMode); safecall;
end;
type TSafeModeImpl = class(TInterfacedObject, ISafeMode)
  public procedure SetMode(m: TMode); safecall;
end;
procedure TSafeModeImpl.SetMode(m: TMode);
begin
  WriteLn('Mode:' + Ord(m).ToString);
end;
var obj: ISafeMode;
begin
  obj := TSafeModeImpl.Create;
  obj.SetMode(mOn);
end.
"#);
    assert_eq!(out, vec!["Mode:1"]);
}

#[test]
fn test_safecall_returning_boolean() {
    let out = run_pascal(r#"
program Test;
type ISafeCheck = interface
  ['{88888888-9999-0000-1111-222222222222}']
  function CheckValue(val: Integer): Boolean; safecall;
end;
type TSafeCheckImpl = class(TInterfacedObject, ISafeCheck)
  public function CheckValue(val: Integer): Boolean; safecall;
end;
function TSafeCheckImpl.CheckValue(val: Integer): Boolean;
begin
  Result := val > 50;
end;
var obj: ISafeCheck;
begin
  obj := TSafeCheckImpl.Create;
  WriteLn(obj.CheckValue(75));
  WriteLn(obj.CheckValue(25));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_safecall_returning_real() {
    let out = run_pascal(r#"
program Test;
type ISafeMath = interface
  ['{99999999-0000-1111-2222-333333333333}']
  function SqrtVal(x: Real): Real; safecall;
end;
type TSafeMathImpl = class(TInterfacedObject, ISafeMath)
  public function SqrtVal(x: Real): Real; safecall;
end;
function TSafeMathImpl.SqrtVal(x: Real): Real;
begin
  Result := Sqrt(x);
end;
var obj: ISafeMath;
begin
  obj := TSafeMathImpl.Create;
  WriteLn(obj.SqrtVal(25.0));
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_safecall_default_parameter() {
    let out = run_pascal(r#"
program Test;
type ISafeDefault = interface
  ['{00000000-1111-2222-3333-444444444444}']
  procedure Execute(prefix: String = 'DefaultPrefix'); safecall;
end;
type TSafeDefaultImpl = class(TInterfacedObject, ISafeDefault)
  public procedure Execute(prefix: String); safecall;
end;
procedure TSafeDefaultImpl.Execute(prefix: String);
begin
  WriteLn('EXEC:' + prefix);
end;
var obj: ISafeDefault;
begin
  obj := TSafeDefaultImpl.Create;
  obj.Execute;
  obj.Execute('CustomPrefix');
end.
"#);
    assert_eq!(out, vec!["EXEC:DefaultPrefix", "EXEC:CustomPrefix"]);
}

#[test]
fn test_safecall_multiple_methods_in_interface() {
    let out = run_pascal(r#"
program Test;
type ISafeMulti = interface
  ['{10101010-2020-3030-4040-505050505050}']
  procedure Step1; safecall;
  procedure Step2; safecall;
end;
type TSafeMultiImpl = class(TInterfacedObject, ISafeMulti)
  public procedure Step1; safecall; procedure Step2; safecall;
end;
procedure TSafeMultiImpl.Step1; begin WriteLn('Step1Executed'); end;
procedure TSafeMultiImpl.Step2; begin WriteLn('Step2Executed'); end;
var obj: ISafeMulti;
begin
  obj := TSafeMultiImpl.Create;
  obj.Step1;
  obj.Step2;
end.
"#);
    assert_eq!(out, vec!["Step1Executed", "Step2Executed"]);
}

#[test]
fn test_safecall_safecallerror_virtual_override() {
    let out = run_pascal(r#"
program Test;
type TSafeCustomObj = class(TInterfacedObject)
  public function SafeCallError(errorCode: HRESULT; ErrorAddr: Pointer): HRESULT; override;
end;
function TSafeCustomObj.SafeCallError(errorCode: HRESULT; ErrorAddr: Pointer): HRESULT;
begin
  WriteLn('SafeCallErrorOverridden');
  Result := inherited SafeCallError(errorCode, ErrorAddr);
end;
var obj: TSafeCustomObj;
begin
  obj := TSafeCustomObj.Create;
  WriteLn(Assigned(obj));
  obj.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_safecall_standalone_procedure_declaration() {
    let out = run_pascal(r#"
program Test;
procedure StandaloneSafeProc(msg: String); safecall;
begin
  WriteLn('StandaloneSafecall:' + msg);
end;
begin
  StandaloneSafeProc('Hello');
end.
"#);
    assert_eq!(out, vec!["StandaloneSafecall:Hello"]);
}

#[test]
fn test_safecall_standalone_function_declaration() {
    let out = run_pascal(r#"
program Test;
function StandaloneSafeFunc(x, y: Integer): Integer; safecall;
begin
  Result := x * y;
end;
begin
  WriteLn(StandaloneSafeFunc(6, 7));
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_safecall_in_subclass_virtual_override() {
    let out = run_pascal(r#"
program Test;
type TBaseSafe = class(TInterfacedObject)
  public procedure DoWork; virtual; safecall;
end;
type TSubSafe = class(TBaseSafe)
  public procedure DoWork; override; safecall;
end;
procedure TBaseSafe.DoWork; begin WriteLn('BaseSafe'); end;
procedure TSubSafe.DoWork; begin WriteLn('SubSafe'); end;

var b: TBaseSafe;
begin
  b := TSubSafe.Create;
  b.DoWork;
  b.Free;
end.
"#);
    assert_eq!(out, vec!["SubSafe"]);
}

#[test]
fn test_safecall_array_parameter() {
    let out = run_pascal(r#"
program Test;
type TArr = array[1..3] of Integer;
type ISafeArrHandler = interface
  ['{20202020-3030-4040-5050-606060606060}']
  procedure ProcessArr(const arr: TArr); safecall;
end;
type TSafeArrImpl = class(TInterfacedObject, ISafeArrHandler)
  public procedure ProcessArr(const arr: TArr); safecall;
end;
procedure TSafeArrImpl.ProcessArr(const arr: TArr);
begin
  WriteLn(arr[1] + arr[2] + arr[3]);
end;
var obj: ISafeArrHandler; a: TArr;
begin
  a[1] := 10; a[2] := 20; a[3] := 30;
  obj := TSafeArrImpl.Create;
  obj.ProcessArr(a);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_safecall_nested_interface_call() {
    let out = run_pascal(r#"
program Test;
type ISafeCore = interface
  ['{30303030-4040-5050-6060-707070707070}']
  procedure RunCore; safecall;
end;
type TSafeCoreImpl = class(TInterfacedObject, ISafeCore)
  public procedure RunCore; safecall;
end;
procedure TSafeCoreImpl.RunCore; begin WriteLn('NestedSafeCoreRun'); end;

var core: ISafeCore;
begin
  core := TSafeCoreImpl.Create;
  core.RunCore;
end.
"#);
    assert_eq!(out, vec!["NestedSafeCoreRun"]);
}

#[test]
fn test_safecall_div_by_zero_exception() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ISafeDiv = interface
  ['{40404040-5050-6060-7070-808080808080}']
  function SafeDiv(a, b: Integer): Integer; safecall;
end;
type TSafeDivImpl = class(TInterfacedObject, ISafeDiv)
  public function SafeDiv(a, b: Integer): Integer; safecall;
end;
function TSafeDivImpl.SafeDiv(a, b: Integer): Integer;
begin
  Result := a div b;
end;

var obj: ISafeDiv;
begin
  obj := TSafeDivImpl.Create;
  try
    obj.SafeDiv(10, 0);
  except
    on E: EDivByZero do WriteLn('CaughtSafecallDivByZero');
  end;
end.
"#);
    assert_eq!(out, vec!["CaughtSafecallDivByZero"]);
}

#[test]
fn test_safecall_const_parameter() {
    let out = run_pascal(r#"
program Test;
type ISafeConst = interface
  ['{50505050-6060-7070-8080-909090909090}']
  procedure LogConst(const msg: String); safecall;
end;
type TSafeConstImpl = class(TInterfacedObject, ISafeConst)
  public procedure LogConst(const msg: String); safecall;
end;
procedure TSafeConstImpl.LogConst(const msg: String);
begin
  WriteLn('CONST:' + msg);
end;
var obj: ISafeConst;
begin
  obj := TSafeConstImpl.Create;
  obj.LogConst('ImmutableSafecallText');
end.
"#);
    assert_eq!(out, vec!["CONST:ImmutableSafecallText"]);
}

#[test]
fn test_safecall_loop_execution() {
    let out = run_pascal(r#"
program Test;
type ISafeLoop = interface
  ['{60606060-7070-8080-9090-000000000000}']
  procedure Ping(idx: Integer); safecall;
end;
type TSafeLoopImpl = class(TInterfacedObject, ISafeLoop)
  public procedure Ping(idx: Integer); safecall;
end;
procedure TSafeLoopImpl.Ping(idx: Integer);
begin
  WriteLn('Ping:' + idx.ToString);
end;
var obj: ISafeLoop; i: Integer;
begin
  obj := TSafeLoopImpl.Create;
  for i := 1 to 3 do
    obj.Ping(i);
end.
"#);
    assert_eq!(out, vec!["Ping:1", "Ping:2", "Ping:3"]);
}
