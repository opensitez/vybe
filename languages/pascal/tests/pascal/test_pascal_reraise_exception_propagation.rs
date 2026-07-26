use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 55: Exception Re-raising & Call Stack Propagation (raise;)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_reraise_untyped_except_block() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('UntypedReraise');
    except
      WriteLn('UntypedIntercepted');
      raise;
    end;
  except
    on E: Exception do WriteLn('OuterCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["UntypedIntercepted", "OuterCaught:UntypedReraise"]
    );
}

#[test]
fn test_reraise_typed_on_e_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise EInvalidArgument.Create('BadArg');
    except
      on E: EInvalidArgument do
      begin
        WriteLn('TypedIntercepted:' + E.Message);
        raise;
      end;
    end;
  except
    on E: EInvalidArgument do WriteLn('OuterTypedCaught:' + E.ClassName);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "TypedIntercepted:BadArg",
            "OuterTypedCaught:EInvalidArgument"
        ]
    );
}

#[test]
fn test_reraise_preserves_exact_subclass_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomChild = class(Exception);
begin
  try
    try
      raise ECustomChild.Create('ChildMessage');
    except
      on E: Exception do
      begin
        WriteLn('BaseIntercept:' + E.ClassName);
        raise;
      end;
    end;
  except
    on E: ECustomChild do WriteLn('OuterChildCaught:' + E.ClassName);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "BaseIntercept:ECustomChild",
            "OuterChildCaught:ECustomChild"
        ]
    );
}

#[test]
fn test_reraise_in_nested_procedure_call() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure MiddleLogger;
begin
  try
    raise Exception.Create('DeepError');
  except
    on E: Exception do
    begin
      WriteLn('LoggerLogged:' + E.Message);
      raise;
    end;
  end;
end;
begin
  try
    MiddleLogger;
  except
    on E: Exception do WriteLn('TopHandler:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["LoggerLogged:DeepError", "TopHandler:DeepError"]);
}

#[test]
fn test_reraise_after_resource_cleanup() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils, Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  try
    try
      sl.Add('Line1');
      raise Exception.Create('FailWithResource');
    except
      WriteLn('CleaningUpResource');
      sl.Free;
      sl := nil;
      raise;
    end;
  except
    on E: Exception do WriteLn('OuterCaughtResourceCleaned:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "CleaningUpResource",
            "OuterCaughtResourceCleaned:FailWithResource"
        ]
    );
}

#[test]
fn test_reraise_conditional() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure Process(shouldReraise: Boolean);
begin
  try
    raise Exception.Create('ConditionalFail');
  except
    on E: Exception do
    begin
      WriteLn('CaughtInProc');
      if shouldReraise then raise;
    end;
  end;
end;
begin
  try
    Process(False);
    WriteLn('ContinuedWhenNoReraise');
    Process(True);
  except
    on E: Exception do WriteLn('CaughtTopAfterReraise');
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "CaughtInProc",
            "ContinuedWhenNoReraise",
            "CaughtInProc",
            "CaughtTopAfterReraise"
        ]
    );
}

#[test]
fn test_reraise_inside_loop_iteration() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var i: Integer;
begin
  for i := 1 to 2 do
  begin
    try
      try
        if i = 2 then raise Exception.Create('LoopErr2');
        WriteLn('LoopOK:' + i.ToString);
      except
        on E: Exception do
        begin
          WriteLn('LoopIntercept:' + i.ToString);
          raise;
        end;
      end;
    except
      on E: Exception do WriteLn('LoopOuterCaught:' + i.ToString);
    end;
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["LoopOK:1", "LoopIntercept:2", "LoopOuterCaught:2"]
    );
}

#[test]
fn test_reraise_updates_global_error_count() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var errorCount: Integer;
procedure AuditError;
begin
  try
    raise Exception.Create('AuditFail');
  except
    on E: Exception do
    begin
      Inc(errorCount);
      raise;
    end;
  end;
end;
begin
  errorCount := 0;
  try
    AuditError;
  except
    on E: Exception do WriteLn('ErrorCount:' + errorCount.ToString);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ErrorCount:1"]);
}

#[test]
fn test_reraise_in_class_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TWorker = class
  public procedure DoTask;
end;
procedure TWorker.DoTask;
begin
  try
    raise Exception.Create('TaskFailed');
  except
    on E: Exception do
    begin
      WriteLn('MethodLogged:' + E.Message);
      raise;
    end;
  end;
end;
var w: TWorker;
begin
  w := TWorker.Create;
  try
    w.DoTask;
  except
    on E: Exception do WriteLn('CallerCaught:' + E.Message);
  end;
  w.Free;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["MethodLogged:TaskFailed", "CallerCaught:TaskFailed"]
    );
}

#[test]
fn test_reraise_in_record_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TRec = record
  procedure Exec;
end;
procedure TRec.Exec;
begin
  try
    raise Exception.Create('RecErr');
  except
    on E: Exception do
    begin
      WriteLn('RecLogged:' + E.Message);
      raise;
    end;
  end;
end;
var r: TRec;
begin
  try
    r.Exec;
  except
    on E: Exception do WriteLn('OuterRecCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RecLogged:RecErr", "OuterRecCaught:RecErr"]);
}

#[test]
fn test_reraise_multiple_chained_levels() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure P1; begin raise Exception.Create('ChainedErr'); end;
procedure P2;
begin
  try P1; except WriteLn('P2Reraising'); raise; end;
end;
procedure P3;
begin
  try P2; except WriteLn('P3Reraising'); raise; end;
end;
begin
  try
    P3;
  except
    on E: Exception do WriteLn('MainCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["P2Reraising", "P3Reraising", "MainCaught:ChainedErr"]
    );
}

#[test]
fn test_reraise_with_finally_block_cleanup() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      try
        raise Exception.Create('WithFinally');
      except
        WriteLn('ExceptBlock');
        raise;
      end;
    finally
      WriteLn('FinallyBlock');
    end;
  except
    on E: Exception do WriteLn('OuterBlock:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["ExceptBlock", "FinallyBlock", "OuterBlock:WithFinally"]
    );
}

#[test]
fn test_reraise_custom_exception_with_extra_fields() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomData = class(Exception)
  public Tag: Integer;
  constructor CreateTag(T: Integer; const msg: String);
end;
constructor ECustomData.CreateTag(T: Integer; const msg: String);
begin
  inherited Create(msg); Tag := T;
end;
begin
  try
    try
      raise ECustomData.CreateTag(999, 'TaggedErr');
    except
      on E: ECustomData do
      begin
        WriteLn('InnerTag:' + E.Tag.ToString);
        raise;
      end;
    end;
  except
    on E: ECustomData do WriteLn('OuterTag:' + E.Tag.ToString);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["InnerTag:999", "OuterTag:999"]);
}

#[test]
fn test_reraise_does_not_alter_message() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var origMsg, caughtMsg: String;
begin
  origMsg := 'ExactOriginalMessage_12345';
  try
    try
      raise Exception.Create(origMsg);
    except
      raise;
    end;
  except
    on E: Exception do caughtMsg := E.Message;
  end;
  WriteLn(origMsg = caughtMsg);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reraise_in_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TChildObj = class
  constructor Create;
end;
type TParentObj = class
  private FChild: TChildObj;
  public constructor Create;
end;
constructor TChildObj.Create; begin raise Exception.Create('ChildCtorErr'); end;
constructor TParentObj.Create;
begin
  try
    FChild := TChildObj.Create;
  except
    WriteLn('ParentCtorLogging');
    raise;
  end;
end;
var p: TParentObj;
begin
  try
    p := TParentObj.Create;
  except
    on E: Exception do WriteLn('CallerCtorCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["ParentCtorLogging", "CallerCtorCaught:ChildCtorErr"]
    );
}

#[test]
fn test_reraise_in_destructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TFailDtor = class
  destructor Destroy; override;
end;
destructor TFailDtor.Destroy;
begin
  try
    raise Exception.Create('DtorErr');
  except
    WriteLn('DtorLogging');
    raise;
  end;
  inherited Destroy;
end;
var obj: TFailDtor;
begin
  obj := TFailDtor.Create;
  try
    obj.Free;
  except
    on E: Exception do WriteLn('FreeCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["DtorLogging", "FreeCaught:DtorErr"]);
}

#[test]
fn test_reraise_in_class_helper() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TIntHelper = record helper for Integer
  public procedure TestHelper;
end;
procedure TIntHelper.TestHelper;
begin
  try
    raise Exception.Create('HelperErr');
  except
    WriteLn('HelperLogged');
    raise;
  end;
end;
var val: Integer;
begin
  val := 10;
  try
    val.TestHelper;
  except
    on E: Exception do WriteLn('CallerHelperCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["HelperLogged", "CallerHelperCaught:HelperErr"]);
}

#[test]
fn test_reraise_preserves_error_code_in_sys_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise EDivByZero.Create('DivByZeroErr');
    except
      on E: EDivByZero do
      begin
        WriteLn('InnerDivByZero');
        raise;
      end;
    end;
  except
    on E: EDivByZero do WriteLn('OuterDivByZeroConfirmed');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["InnerDivByZero", "OuterDivByZeroConfirmed"]);
}

#[test]
fn test_reraise_in_interface_method_implementation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type IWork = interface
  ['{12341234-1234-1234-1234-123412341234}']
  procedure DoWork;
end;
type TWorkImpl = class(TInterfacedObject, IWork)
  public procedure DoWork;
end;
procedure TWorkImpl.DoWork;
begin
  try
    raise Exception.Create('IntfWorkErr');
  except
    WriteLn('IntfLogged');
    raise;
  end;
end;
var w: IWork;
begin
  w := TWorkImpl.Create;
  try
    w.DoWork;
  except
    on E: Exception do WriteLn('IntfCallerCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["IntfLogged", "IntfCallerCaught:IntfWorkErr"]);
}

#[test]
fn test_reraise_in_function_return() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
function GetVal: Integer;
begin
  try
    raise Exception.Create('FuncValErr');
  except
    WriteLn('FuncLogged');
    raise;
  end;
end;
begin
  try
    GetVal;
  except
    on E: Exception do WriteLn('FuncCallerCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["FuncLogged", "FuncCallerCaught:FuncValErr"]);
}
