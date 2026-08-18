use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 53: Nested Exception Handling & Stack Unwinding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_nested_try_except_inner_catches_first() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('InnerError');
    except
      on E: Exception do WriteLn('InnerCaught:' + E.Message);
    end;
  except
    on E: Exception do WriteLn('OuterCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["InnerCaught:InnerError"]);
}

#[test]
fn test_nested_try_except_reraise_to_outer() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('OriginalError');
    except
      on E: Exception do
      begin
        WriteLn('InnerLog:' + E.Message);
        raise;
      end;
    end;
  except
    on E: Exception do WriteLn('OuterCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InnerLog:OriginalError", "OuterCaught:OriginalError"]
    );
}

#[test]
fn test_nested_try_except_wrap_new_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise EDivByZero.Create('DivisionByZero');
    except
      on E: EDivByZero do
        raise EInvalidArgument.Create('Wrapped:' + E.Message);
    end;
  except
    on E: EInvalidArgument do WriteLn('OuterCaughtWrapped:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["OuterCaughtWrapped:Wrapped:DivisionByZero"]);
}

#[test]
fn test_exception_in_finally_block() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      WriteLn('InTryBlock');
    finally
      raise Exception.Create('FinallyFailed');
    end;
  except
    on E: Exception do WriteLn('OuterCaughtFinallyErr:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InTryBlock", "OuterCaughtFinallyErr:FinallyFailed"]
    );
}

#[test]
fn test_stack_unwinding_across_three_procedures() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure Level3;
begin
  raise Exception.Create('DeepFailure');
end;
procedure Level2;
begin
  Level3;
end;
procedure Level1;
begin
  Level2;
end;
begin
  try
    Level1;
  except
    on E: Exception do WriteLn('UnwoundToTop:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["UnwoundToTop:DeepFailure"]);
}

#[test]
fn test_try_finally_inside_try_except() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('WorkError');
    finally
      WriteLn('CleanupExecuted');
    end;
  except
    on E: Exception do WriteLn('HandlerExecuted:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CleanupExecuted", "HandlerExecuted:WorkError"]);
}

#[test]
fn test_try_except_inside_try_finally() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('HandledInner');
    except
      on E: Exception do WriteLn('InnerHandled:' + E.Message);
    end;
  finally
    WriteLn('OuterFinallyExecuted');
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InnerHandled:HandledInner", "OuterFinallyExecuted"]
    );
}

#[test]
fn test_selective_inner_catch_unmatched_bubbles_up() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise EAccessViolation.Create('AVError');
    except
      on E: EDivByZero do WriteLn('InnerDivZeroNotTriggered');
    end;
  except
    on E: EAccessViolation do WriteLn('OuterCaughtAV:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["OuterCaughtAV:AVError"]);
}

#[test]
fn test_recursive_unwinding_with_finally() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure Recurse(depth: Integer);
begin
  try
    if depth = 3 then raise Exception.Create('Depth3Err')
    else Recurse(depth + 1);
  finally
    WriteLn('UnwindingDepth:' + depth.ToString);
  end;
end;
begin
  try
    Recurse(1);
  except
    on E: Exception do WriteLn('TopCaught:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "UnwindingDepth:3",
            "UnwindingDepth:2",
            "UnwindingDepth:1",
            "TopCaught:Depth3Err"
        ]
    );
}

#[test]
fn test_nested_try_except_in_loop() {
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
        if i = 1 then raise Exception.Create('Err1')
        else raise EDivByZero.Create('Err2');
      except
        on E: EDivByZero do WriteLn('InnerCaughtDivZero');
      end;
    except
      on E: Exception do WriteLn('OuterCaughtErr:' + E.Message);
    end;
  end;
end.
"#,
    );
    assert_eq!(out, vec!["OuterCaughtErr:Err1", "InnerCaughtDivZero"]);
}

#[test]
fn test_exception_during_except_handler_execution() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('FirstErr');
    except
      on E: Exception do
      begin
        WriteLn('FirstCaught:' + E.Message);
        raise Exception.Create('SecondErrFromExcept');
      end;
    end;
  except
    on E: Exception do WriteLn('OuterCaughtSecond:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec![
            "FirstCaught:FirstErr",
            "OuterCaughtSecond:SecondErrFromExcept"
        ]
    );
}

#[test]
fn test_nested_exceptions_in_class_hierarchy() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TSubProcessor = class
  public procedure Process;
end;
type TMainProcessor = class
  private FSub: TSubProcessor;
  public constructor Create; procedure Execute;
end;
procedure TSubProcessor.Process; begin raise Exception.Create('SubFail'); end;
constructor TMainProcessor.Create; begin FSub := TSubProcessor.Create; end;
procedure TMainProcessor.Execute;
begin
  try
    FSub.Process;
  finally
    FSub.Free;
    WriteLn('SubFreedInFinally');
  end;
end;
var mainProc: TMainProcessor;
begin
  mainProc := TMainProcessor.Create;
  try
    mainProc.Execute;
  except
    on E: Exception do WriteLn('MainCaught:' + E.Message);
  end;
  mainProc.Free;
end.
"#,
    );
    assert_eq!(out, vec!["SubFreedInFinally", "MainCaught:SubFail"]);
}

#[test]
fn test_nested_exception_in_property_getter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TTestProp = class
  private function GetVal: Integer;
  public property Val: Integer read GetVal;
end;
function TTestProp.GetVal: Integer;
begin
  raise Exception.Create('GetterError');
end;
var t: TTestProp;
begin
  t := TTestProp.Create;
  try
    try
      WriteLn(t.Val);
    except
      on E: Exception do WriteLn('CaughtGetter:' + E.Message);
    end;
  finally
    t.Free;
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CaughtGetter:GetterError"]);
}

#[test]
fn test_nested_exception_in_property_setter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TTestProp = class
  private procedure SetVal(v: Integer);
  public property Val: Integer write SetVal;
end;
procedure TTestProp.SetVal(v: Integer);
begin
  if v < 0 then raise Exception.Create('NegativeValueNotAllowed');
end;
var t: TTestProp;
begin
  t := TTestProp.Create;
  try
    t.Val := -5;
  except
    on E: Exception do WriteLn('CaughtSetter:' + E.Message);
  end;
  t.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CaughtSetter:NegativeValueNotAllowed"]);
}

#[test]
fn test_nested_try_except_in_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TSubObj = class
  constructor Create;
end;
type TParentObj = class
  public Sub: TSubObj;
  constructor Create;
end;
constructor TSubObj.Create; begin raise Exception.Create('SubCtorErr'); end;
constructor TParentObj.Create;
begin
  try
    Sub := TSubObj.Create;
  except
    on E: Exception do
    begin
      WriteLn('ParentCtorHandled:' + E.Message);
      Sub := nil;
    end;
  end;
end;
var p: TParentObj;
begin
  p := TParentObj.Create;
  WriteLn(p.Sub = nil);
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ParentCtorHandled:SubCtorErr", "TRUE"]);
}

#[test]
fn test_nested_exception_preserves_class_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise ERangeError.Create('OutOfRange');
    except
      raise;
    end;
  except
    on E: Exception do WriteLn(E.ClassName + ':' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ERangeError:OutOfRange"]);
}

#[test]
fn test_nested_finally_with_reraised_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      try
        raise Exception.Create('DeepErr');
      finally
        WriteLn('Fin1');
      end;
    finally
      WriteLn('Fin2');
    end;
  except
    on E: Exception do WriteLn('CaughtAtTop:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Fin1", "Fin2", "CaughtAtTop:DeepErr"]);
}

#[test]
fn test_nested_exception_handling_in_record() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TRec = record
  procedure Run;
end;
procedure TRec.Run;
begin
  try
    raise Exception.Create('RecRunErr');
  except
    on E: Exception do WriteLn('RecHandled:' + E.Message);
  end;
end;
var r: TRec;
begin
  r.Run;
end.
"#,
    );
    assert_eq!(out, vec!["RecHandled:RecRunErr"]);
}

#[test]
fn test_nested_exception_with_array_processing() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure ProcessArray(const arr: array of Integer);
var i: Integer;
begin
  for i := Low(arr) to High(arr) do
  begin
    try
      if arr[i] = 0 then raise EDivByZero.Create('ZeroElement');
      WriteLn('Val:' + arr[i].ToString);
    except
      on E: EDivByZero do WriteLn('SkippedZeroAtIndex:' + i.ToString);
    end;
  end;
end;
begin
  ProcessArray([10, 0, 30]);
end.
"#,
    );
    assert_eq!(out, vec!["Val:10", "SkippedZeroAtIndex:1", "Val:30"]);
}

#[test]
fn test_three_level_nested_try_except() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      try
        raise Exception.Create('Level3Err');
      except
        on E: EDivByZero do WriteLn('Level3Handled');
      end;
    except
      on E: EArgumentException do WriteLn('Level2Handled');
    end;
  except
    on E: Exception do WriteLn('Level1Handled:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Level1Handled:Level3Err"]);
}
