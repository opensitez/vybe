/// Tests for advanced exception handling in Pascal/Delphi:
/// re-raise patterns, exception message access, exception class hierarchy,
/// multiple on-clauses, exception in loops, nested exception handlers.
use super::helpers::run_pascal;

// ===================================================================
// RE-RAISE (bare RAISE in except block)
// ===================================================================

#[test]
fn reraise_caught_by_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('inner error');
    except
      on E: Exception do
      begin
        WriteLn('inner caught: ' + E.Message);
        raise;
      end;
    end;
  except
    on E: Exception do
      WriteLn('outer caught: ' + E.Message);
  end;
end."#
        ),
        &["inner caught: inner error", "outer caught: inner error"]
    );
}

// ===================================================================
// EXCEPTION MESSAGE PROPERTY
// ===================================================================

#[test]
fn exception_message_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('something went wrong');
  except
    on E: Exception do
      WriteLn(E.Message);
  end;
end."#
        ),
        &["something went wrong"]
    );
}

#[test]
fn exception_message_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SafeDivide(a, b: Integer): Integer;
begin
  if b = 0 then
    raise Exception.Create('division by zero');
  Result := a div b;
end;
begin
  try
    WriteLn(SafeDivide(10, 2));
    WriteLn(SafeDivide(5, 0));
  except
    on E: Exception do
      WriteLn('Error: ' + E.Message);
  end;
end."#
        ),
        &["5", "Error: division by zero"]
    );
}

// ===================================================================
// MULTIPLE ON-CLAUSES (EXCEPTION CLASS HIERARCHY)
// ===================================================================

#[test]
fn multiple_on_clauses_order() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  EMyError = class(Exception);
  ESpecificError = class(EMyError);
begin
  try
    raise ESpecificError.Create('specific');
  except
    on E: ESpecificError do WriteLn('specific handler');
    on E: EMyError do WriteLn('my error handler');
    on E: Exception do WriteLn('base handler');
  end;
end."#
        ),
        &["specific handler"]
    );
}

#[test]
fn exception_base_catches_derived() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  EDerivedError = class(Exception);
begin
  try
    raise EDerivedError.Create('derived');
  except
    on E: Exception do
      WriteLn('caught: ' + E.Message);
  end;
end."#
        ),
        &["caught: derived"]
    );
}

// ===================================================================
// EXCEPTION IN LOOPS
// ===================================================================

#[test]
fn exception_continues_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  for i := 1 to 3 do
  begin
    try
      if i = 2 then raise Exception.Create('skip');
      WriteLn(i);
    except
      on E: Exception do WriteLn('skip');
    end;
  end;
end."#
        ),
        &["1", "skip", "3"]
    );
}

#[test]
fn exception_counter_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, errCount: Integer;
begin
  errCount := 0;
  for i := 1 to 5 do
  begin
    try
      if Odd(i) then raise Exception.Create('odd');
    except
      Inc(errCount);
    end;
  end;
  WriteLn(errCount);
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// TRY-FINALLY ORDERING
// ===================================================================

#[test]
fn finally_before_except_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('test');
    finally
      WriteLn('finally runs');
    end;
  except
    WriteLn('except runs');
  end;
end."#
        ),
        &["finally runs", "except runs"]
    );
}

#[test]
fn finally_always_runs_on_exit() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure DoWork;
begin
  try
    WriteLn('working');
    Exit;
  finally
    WriteLn('cleanup');
  end;
end;
begin
  DoWork;
  WriteLn('done');
end."#
        ),
        &["working", "cleanup", "done"]
    );
}

// ===================================================================
// CUSTOM EXCEPTION WITH EXTRA DATA
// ===================================================================

#[test]
fn custom_exception_with_code() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAppError = class(Exception)
  public
    Code: Integer;
    constructor CreateWithCode(aMsg: String; aCode: Integer);
  end;
constructor TAppError.CreateWithCode(aMsg: String; aCode: Integer);
begin
  inherited Create(aMsg);
  Code := aCode;
end;
begin
  try
    raise TAppError.CreateWithCode('not found', 404);
  except
    on E: TAppError do
      WriteLn(IntToStr(E.Code) + ': ' + E.Message);
  end;
end."#
        ),
        &["404: not found"]
    );
}

#[test]
fn exception_with_classname() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  EValidationError = class(Exception);
begin
  try
    raise EValidationError.Create('invalid input');
  except
    on E: EValidationError do
      WriteLn('Validation: ' + E.Message);
    on E: Exception do
      WriteLn('Generic: ' + E.Message);
  end;
end."#
        ),
        &["Validation: invalid input"]
    );
}

// ===================================================================
// EXCEPTION IN NESTED PROCEDURE
// ===================================================================

#[test]
fn exception_propagates_from_nested() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Inner;
begin
  raise Exception.Create('from inner');
end;
procedure Outer;
begin
  Inner;
end;
begin
  try
    Outer;
  except
    on E: Exception do
      WriteLn('caught: ' + E.Message);
  end;
end."#
        ),
        &["caught: from inner"]
    );
}

// ===================================================================
// EXCEPTION WITH TRY-FINALLY AND RESOURCE
// ===================================================================

#[test]
fn resource_cleanup_on_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
var acquired: Boolean;
begin
  acquired := False;
  try
    acquired := True;
    WriteLn('acquired');
    raise Exception.Create('fail');
  except
    on E: Exception do
      WriteLn('handled: ' + E.Message);
  end;
  if acquired then
    WriteLn('cleanup needed: ' + BoolToStr(acquired, True));
end."#
        ),
        &["acquired", "handled: fail", "cleanup needed: True"]
    );
}

// ===================================================================
// EXCEPTION CLASS HIERARCHY
// ===================================================================

#[test]
fn three_level_exception_hierarchy() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  ELevel1 = class(Exception);
  ELevel2 = class(ELevel1);
  ELevel3 = class(ELevel2);
begin
  try
    raise ELevel3.Create('deep');
  except
    on E: ELevel1 do
      WriteLn('caught at level 1: ' + E.Message);
  end;
end."#
        ),
        &["caught at level 1: deep"]
    );
}

// ===================================================================
// EXCEPTION MESSAGE FORMATTING
// ===================================================================

#[test]
fn exception_formatted_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure CheckRange(val, lo, hi: Integer);
begin
  if (val < lo) or (val > hi) then
    raise Exception.Create(
      Format('Value %d out of range [%d..%d]', [val, lo, hi]));
end;
begin
  try
    CheckRange(50, 1, 10);
  except
    on E: Exception do
      WriteLn(E.Message);
  end;
end."#
        ),
        &["Value 50 out of range [1..10]"]
    );
}

// ===================================================================
// MULTIPLE EXCEPTIONS IN SEQUENCE
// ===================================================================

#[test]
fn multiple_separate_exceptions() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
    total: Integer;
begin
  total := 0;
  for i := 1 to 4 do
  begin
    try
      case i of
        1: raise Exception.Create('one');
        3: raise Exception.Create('three');
      end;
      total := total + i;
    except
      on E: Exception do Inc(total, 100);
    end;
  end;
  WriteLn(total);
end."#
        ),
        &["206"]
    );
}

// ===================================================================
// EXCEPTION WITH ELSE CLAUSE
// ===================================================================

#[test]
fn try_except_no_raise_runs_normally() {
    assert_eq!(
        run_pascal(
            r#"program T;
function TryParse(s: String): Integer;
begin
  try
    Result := StrToInt(s);
  except
    Result := -1;
  end;
end;
begin
  WriteLn(TryParse('42'));
  WriteLn(TryParse('bad'));
end."#
        ),
        &["42", "-1"]
    );
}

// ===================================================================
// FINALLY WITH EXCEPTION SWALLOWED
// ===================================================================

#[test]
fn finally_with_swallowed_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
var log: String;
begin
  log := '';
  try
    try
      log := log + 'try ';
      raise Exception.Create('x');
    except
      log := log + 'except ';
    end;
  finally
    log := log + 'finally';
  end;
  WriteLn(log);
end."#
        ),
        &["try except finally"]
    );
}
