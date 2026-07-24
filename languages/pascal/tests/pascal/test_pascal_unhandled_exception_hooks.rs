use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 60: Unhandled Exception Hooks (ExceptProc)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_exceptproc_hook_registration() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var oldProc: TExceptProc;

procedure CustomExceptProc(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('HookCaught:' + Exception(ExceptObj).Message);
end;

begin
  oldProc := ExceptProc;
  ExceptProc := CustomExceptProc;

  raise Exception.Create('UnhandledError');

  ExceptProc := oldProc;
end.
"#);
    assert_eq!(out, vec!["HookCaught:UnhandledError"]);
}

#[test]
fn test_exceptproc_classname_logging() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure LogExceptClass(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('HookClass:' + ExceptObj.ClassName);
end;
begin
  ExceptProc := LogExceptClass;
  raise EInvalidArgument.Create('BadArg');
end.
"#);
    assert_eq!(out, vec!["HookClass:EInvalidArgument"]);
}

#[test]
fn test_exceptproc_exceptaddr_not_nil() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure CheckAddrHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn(ExceptAddr <> nil);
end;
begin
  ExceptProc := CheckAddrHook;
  raise Exception.Create('AddrCheck');
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptproc_custom_exception_casting() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ECustomHookErr = class(Exception)
  public Code: Integer;
  constructor CreateCode(C: Integer; const msg: String);
end;
constructor ECustomHookErr.CreateCode(C: Integer; const msg: String);
begin
  inherited Create(msg); Code := C;
end;

procedure CustomHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  if ExceptObj is ECustomHookErr then
    WriteLn('Code:' + ECustomHookErr(ExceptObj).Code.ToString);
end;

begin
  ExceptProc := CustomHook;
  raise ECustomHookErr.CreateCode(500, 'ServerErr');
end.
"#);
    assert_eq!(out, vec!["Code:500"]);
}

#[test]
fn test_exceptproc_global_counter_increment() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var unhandledCount: Integer;
procedure CountHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  Inc(unhandledCount);
  WriteLn('UnhandledCount:' + unhandledCount.ToString);
end;
begin
  unhandledCount := 0;
  ExceptProc := CountHook;
  raise Exception.Create('CountTest');
end.
"#);
    assert_eq!(out, vec!["UnhandledCount:1"]);
}

#[test]
fn test_exceptproc_restoration_to_previous() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var oldProc: TExceptProc;
procedure DummyHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('DummyHookHandled');
end;
begin
  oldProc := ExceptProc;
  ExceptProc := DummyHook;
  ExceptProc := oldProc;
  WriteLn(ExceptProc = oldProc);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptproc_edivbyzero() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure DivZeroHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  if ExceptObj is EDivByZero then
    WriteLn('DivZeroHookCaught');
end;
var a, b: Integer;
begin
  ExceptProc := DivZeroHook;
  a := 10; b := 0;
  a := a div b;
end.
"#);
    assert_eq!(out, vec!["DivZeroHookCaught"]);
}

#[test]
fn test_exceptproc_econvert_error() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure ConvertHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  if ExceptObj is EConvertError then
    WriteLn('ConvertHookCaught:' + Exception(ExceptObj).Message);
end;
begin
  ExceptProc := ConvertHook;
  StrToInt('InvalidInt');
end.
"#);
    assert_eq!(out, vec!["ConvertHookCaught:InvalidInt"]);
}

#[test]
fn test_exceptproc_eaccessviolation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure AVHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  if ExceptObj is EAccessViolation then
    WriteLn('AVHookCaught');
end;
var p: PInteger;
begin
  ExceptProc := AVHook;
  p := nil;
  p^ := 100;
end.
"#);
    assert_eq!(out, vec!["AVHookCaught"]);
}

#[test]
fn test_exceptproc_chained_delegation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var originalHook: TExceptProc;

procedure ChainedHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('ChainedHook1');
  if Assigned(originalHook) then originalHook(ExceptObj, ExceptAddr);
end;

procedure BaseHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('BaseHook2');
end;

begin
  originalHook := BaseHook;
  ExceptProc := ChainedHook;
  raise Exception.Create('ChainedTest');
end.
"#);
    assert_eq!(out, vec!["ChainedHook1", "BaseHook2"]);
}

#[test]
fn test_exceptproc_in_class_method_caller() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TWorker = class
  public procedure Fail;
end;
procedure TWorker.Fail;
begin
  raise Exception.Create('MethodFailed');
end;

procedure MethodHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('MethodHook:' + Exception(ExceptObj).Message);
end;

var w: TWorker;
begin
  ExceptProc := MethodHook;
  w := TWorker.Create;
  w.Fail;
end.
"#);
    assert_eq!(out, vec!["MethodHook:MethodFailed"]);
}

#[test]
fn test_exceptproc_in_record_method_caller() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TRec = record
  procedure Fail;
end;
procedure TRec.Fail;
begin
  raise Exception.Create('RecFailed');
end;

procedure RecHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('RecHook:' + Exception(ExceptObj).Message);
end;

var r: TRec;
begin
  ExceptProc := RecHook;
  r.Fail;
end.
"#);
    assert_eq!(out, vec!["RecHook:RecFailed"]);
}

#[test]
fn test_exceptproc_multi_level_subclass() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type EBase = class(Exception);
type ESub = class(EBase);

procedure HierarchyHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  if ExceptObj is EBase then
    WriteLn('HierarchyMatched:' + ExceptObj.ClassName);
end;

begin
  ExceptProc := HierarchyHook;
  raise ESub.Create('SubClassError');
end.
"#);
    assert_eq!(out, vec!["HierarchyMatched:ESub"]);
}

#[test]
fn test_exceptproc_formatting_hex_addr() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure HexAddrHook(ExceptObj: TObject; ExceptAddr: Pointer);
var hexStr: String;
begin
  hexStr := HexStr(NativeInt(ExceptAddr), 8);
  WriteLn(Length(hexStr) >= 8);
end;
begin
  ExceptProc := HexAddrHook;
  raise Exception.Create('HexAddrTest');
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptproc_function_return_fail() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
function GetVal: Integer;
begin
  raise Exception.Create('FuncFailed');
end;

procedure FuncHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('FuncHook:' + Exception(ExceptObj).Message);
end;

begin
  ExceptProc := FuncHook;
  GetVal;
end.
"#);
    assert_eq!(out, vec!["FuncHook:FuncFailed"]);
}

#[test]
fn test_exceptproc_constructor_fail() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TBadCtor = class
  constructor Create;
end;
constructor TBadCtor.Create;
begin
  raise Exception.Create('CtorFailed');
end;

procedure CtorHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('CtorHook:' + Exception(ExceptObj).Message);
end;

var obj: TBadCtor;
begin
  ExceptProc := CtorHook;
  obj := TBadCtor.Create;
end.
"#);
    assert_eq!(out, vec!["CtorHook:CtorFailed"]);
}

#[test]
fn test_exceptproc_nil_check_assignment() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  ExceptProc := nil;
  WriteLn(ExceptProc = nil);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptproc_with_array_processing() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure ArrayHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('ArrayHook:' + Exception(ExceptObj).Message);
end;
procedure Process(const arr: array of Integer);
begin
  raise Exception.Create('ArrProcessingErr');
end;
begin
  ExceptProc := ArrayHook;
  Process([1, 2, 3]);
end.
"#);
    assert_eq!(out, vec!["ArrayHook:ArrProcessingErr"]);
}

#[test]
fn test_exceptproc_reraised_exception() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure ReraiseHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('ReraiseHook:' + Exception(ExceptObj).Message);
end;
begin
  ExceptProc := ReraiseHook;
  try
    raise Exception.Create('ReraisedUnhandled');
  except
    raise;
  end;
end.
"#);
    assert_eq!(out, vec!["ReraiseHook:ReraisedUnhandled"]);
}

#[test]
fn test_exceptproc_safecall_unhandled() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure SafecallHook(ExceptObj: TObject; ExceptAddr: Pointer);
begin
  WriteLn('SafecallHook:' + Exception(ExceptObj).Message);
end;
procedure DoSafe; safecall;
begin
  raise Exception.Create('SafecallUnhandled');
end;
begin
  ExceptProc := SafecallHook;
  DoSafe;
end.
"#);
    assert_eq!(out, vec!["SafecallHook:SafecallUnhandled"]);
}
