use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 51: Exception Handling (try...except Blocks)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_try_except_basic_catch() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('CustomError');
  except
    WriteLn('Handled');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Handled"]);
}

#[test]
fn test_try_except_on_e_exception_message() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('SpecificMessage');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["SpecificMessage"]);
}

#[test]
fn test_try_except_div_by_zero() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var a, b, c: Integer;
begin
  a := 10; b := 0;
  try
    c := a div b;
    WriteLn(c);
  except
    on E: EDivByZero do WriteLn('DivByZeroCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["DivByZeroCaught"]);
}

#[test]
fn test_try_except_multiple_on_clauses() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure CauseError(kind: Integer);
begin
  try
    if kind = 1 then raise EArgumentException.Create('BadArg')
    else raise EInvalidOp.Create('BadOp');
  except
    on E: EArgumentException do WriteLn('ArgErr:' + E.Message);
    on E: EInvalidOp do WriteLn('OpErr:' + E.Message);
  end;
end;
begin
  CauseError(1);
  CauseError(2);
end.
"#,
    );
    assert_eq!(out, vec!["ArgErr:BadArg", "OpErr:BadOp"]);
}

#[test]
fn test_try_except_else_fallback() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('UnknownException');
  except
    on E: EDivByZero do WriteLn('DivZero');
  else
    WriteLn('FallbackElseBlock');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["FallbackElseBlock"]);
}

#[test]
fn test_try_except_inside_loop() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var i: Integer;
begin
  for i := 1 to 3 do
  begin
    try
      if i = 2 then raise Exception.Create('Err2');
      WriteLn('OK:' + i.ToString);
    except
      on E: Exception do WriteLn('Handled:' + i.ToString);
    end;
  end;
end.
"#,
    );
    assert_eq!(out, vec!["OK:1", "Handled:2", "OK:3"]);
}

#[test]
fn test_try_except_caller_stack_catch() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure Helper;
begin
  raise Exception.Create('NestedFailure');
end;
begin
  try
    Helper;
  except
    on E: Exception do WriteLn('CallerCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CallerCaught:NestedFailure"]);
}

#[test]
fn test_try_except_suppress_and_continue() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('IgnoreMe');
  except
    // Silence exception
  end;
  WriteLn('ExecutionContinued');
end.
"#,
    );
    assert_eq!(out, vec!["ExecutionContinued"]);
}

#[test]
fn test_try_except_return_default_fallback_value() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
function SafeDivide(a, b: Integer): Integer;
begin
  try
    Result := a div b;
  except
    on EDivByZero do Result := -1;
  end;
end;
begin
  WriteLn(SafeDivide(10, 2));
  WriteLn(SafeDivide(10, 0));
end.
"#,
    );
    assert_eq!(out, vec!["5", "-1"]);
}

#[test]
fn test_try_except_in_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TFailObj = class
  constructor Create;
end;
constructor TFailObj.Create;
begin
  raise Exception.Create('ConstructorFailed');
end;
var obj: TFailObj;
begin
  try
    obj := TFailObj.Create;
  except
    on E: Exception do WriteLn('CtorCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CtorCaught:ConstructorFailed"]);
}

#[test]
fn test_try_except_in_record_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TRecWorker = record
  procedure DoWork;
end;
procedure TRecWorker.DoWork;
begin
  raise Exception.Create('RecMethodErr');
end;
var w: TRecWorker;
begin
  try
    w.DoWork;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RecMethodErr"]);
}

#[test]
fn test_try_except_retry_pattern() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var attempts: Integer;
begin
  attempts := 0;
  while attempts < 3 do
  begin
    Inc(attempts);
    try
      if attempts < 3 then raise Exception.Create('TemporaryFail');
      WriteLn('SuccessOnAttempt:' + attempts.ToString);
      Break;
    except
      on E: Exception do WriteLn('RetryingAttempt:' + attempts.ToString);
    end;
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "RetryingAttempt:1",
            "RetryingAttempt:2",
            "SuccessOnAttempt:3"
        ]
    );
}

#[test]
fn test_try_except_no_exception_normal_execution() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    WriteLn('NormalFlow');
  except
    WriteLn('ExceptFlow');
  end;
  WriteLn('Completed');
end.
"#,
    );
    assert_eq!(out, vec!["NormalFlow", "Completed"]);
}

#[test]
fn test_try_except_range_error() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure TriggerRangeError;
var sub: 1..5;
begin
  sub := 10;
  WriteLn(sub);
end;
begin
  try
    TriggerRangeError;
  except
    on E: ERangeError do WriteLn('RangeErrorCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RangeErrorCaught"]);
}

#[test]
fn test_try_except_access_violation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var p: PInteger;
begin
  p := nil;
  try
    p^ := 123;
  except
    on E: EAccessViolation do WriteLn('AVCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["AVCaught"]);
}

#[test]
fn test_try_except_formatting_string_error() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    StrToInt('NotANumber');
  except
    on E: EConvertError do WriteLn('ConvertErrCaught');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ConvertErrCaught"]);
}

#[test]
fn test_try_except_with_local_vars() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var status: String;
begin
  status := 'Initial';
  try
    status := 'InTry';
    raise Exception.Create('Boom');
  except
    on E: Exception do status := 'InExcept:' + E.Message;
  end;
  WriteLn(status);
end.
"#,
    );
    assert_eq!(out, vec!["InExcept:Boom"]);
}

#[test]
fn test_try_except_nested_in_finally() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('InnerError');
    except
      on E: Exception do WriteLn('InnerHandled:' + E.Message);
    end;
  finally
    WriteLn('OuterFinallyExecuted');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["InnerHandled:InnerError", "OuterFinallyExecuted"]);
}

#[test]
fn test_try_except_generic_exception_class_catch() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise EInvalidArgument.Create('BadArgumentValue');
  except
    on E: Exception do WriteLn('BaseExceptionCaught:' + E.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["BaseExceptionCaught:EInvalidArgument"]);
}

#[test]
fn test_try_except_empty_catch_block() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('Ignored');
  except
  end;
  WriteLn('SafelyIgnored');
end.
"#,
    );
    assert_eq!(out, vec!["SafelyIgnored"]);
}
