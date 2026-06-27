/// Exception handling patterns from standard Object Pascal / Delphi.
/// try/except, try/finally, raise, on E: ExceptionType, nested try blocks,
/// re-raise, exception in constructors, custom exception classes.
use super::helpers::run_pascal;

// ===================================================================
// TRY / EXCEPT BASICS
// ===================================================================

#[test]
fn try_except_catches() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('boom');
  except
    WriteLn('caught');
  end;
end."#
        ),
        &["caught"]
    );
}

#[test]
fn try_except_with_on_clause() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('error msg');
  except
    on E: Exception do
      WriteLn(E.Message);
  end;
end."#
        ),
        &["error msg"]
    );
}

#[test]
fn try_except_no_exception_runs_normally() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  try
    x := 42;
    WriteLn(x);
  except
    WriteLn('error');
  end;
end."#
        ),
        &["42"]
    );
}

#[test]
fn try_except_code_after_continues() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('oops');
  except
    WriteLn('handled');
  end;
  WriteLn('continuing');
end."#
        ),
        &["handled", "continuing"]
    );
}

// ===================================================================
// TRY / FINALLY
// ===================================================================

#[test]
fn try_finally_runs_cleanup() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  WriteLn('before');
  try
    WriteLn('inside');
  finally
    WriteLn('cleanup');
  end;
  WriteLn('after');
end."#
        ),
        &["before", "inside", "cleanup", "after"]
    );
}

#[test]
fn try_finally_with_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      WriteLn('start');
      raise Exception.Create('fail');
    finally
      WriteLn('finally runs');
    end;
  except
    WriteLn('caught');
  end;
end."#
        ),
        &["start", "finally runs", "caught"]
    );
}

// ===================================================================
// NESTED TRY BLOCKS
// ===================================================================

#[test]
fn nested_try_except() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('inner');
    except
      WriteLn('caught inner');
    end;
    WriteLn('outer continues');
  except
    WriteLn('caught outer');
  end;
end."#
        ),
        &["caught inner", "outer continues"]
    );
}

#[test]
fn try_except_in_loop() {
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
      WriteLn('error at ' + IntToStr(i));
    end;
  end;
end."#
        ),
        &["1", "error at 2", "3"]
    );
}

// ===================================================================
// TRY/EXCEPT IN FUNCTIONS
// ===================================================================

#[test]
fn try_except_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SafeDiv(a, b: Integer): String;
begin
  try
    if b = 0 then raise Exception.Create('division by zero');
    Result := IntToStr(a div b);
  except
    on E: Exception do
      Result := 'Error: ' + E.Message;
  end;
end;
begin
  WriteLn(SafeDiv(10, 2));
  WriteLn(SafeDiv(10, 0));
end."#
        ),
        &["5", "Error: division by zero"]
    );
}

// ===================================================================
// CUSTOM EXCEPTION CLASSES
// ===================================================================

#[test]
fn custom_exception_class() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  EMyError = class(Exception)
  public
    constructor Create(msg: String);
  end;

constructor EMyError.Create(msg: String);
begin
  inherited Create(msg);
end;

begin
  try
    raise EMyError.Create('custom error');
  except
    on E: EMyError do
      WriteLn('Custom: ' + E.Message);
  end;
end."#
        ),
        &["Custom: custom error"]
    );
}

// ===================================================================
// RAISE IN PROCEDURE
// ===================================================================

#[test]
fn raise_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Validate(x: Integer);
begin
  if x < 0 then raise Exception.Create('negative');
  WriteLn('valid: ' + IntToStr(x));
end;
begin
  try
    Validate(5);
    Validate(-1);
  except
    on E: Exception do WriteLn('Error: ' + E.Message);
  end;
end."#
        ),
        &["valid: 5", "Error: negative"]
    );
}

// ===================================================================
// TRY/FINALLY FOR RESOURCE CLEANUP PATTERN
// ===================================================================

#[test]
fn try_finally_resource_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TResource = class
  public
    FName: String;
    constructor Create(name: String);
    procedure DoWork;
  end;

constructor TResource.Create(name: String);
begin FName := name; WriteLn('Opened ' + name); end;

procedure TResource.DoWork;
begin WriteLn('Working with ' + FName); end;

var r: TResource;
begin
  r := TResource.Create('file.txt');
  try
    r.DoWork;
  finally
    WriteLn('Closing ' + r.FName);
    FreeAndNil(r);
  end;
end."#
        ),
        &[
            "Opened file.txt",
            "Working with file.txt",
            "Closing file.txt"
        ]
    );
}

// -------------------------------------------------------------------
// from test_exceptions_finally_order.rs
// -------------------------------------------------------------------
#[test]
fn finally_runs_after_normal_try_body() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('try');
  finally
    WriteLn('finally');
  end;
  WriteLn('after');
end."#
        ),
        &["try", "finally", "after"]
    );
}

#[test]
fn finally_runs_when_try_raises() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('boom');
    finally
      WriteLn('cleanup');
    end;
  except
    WriteLn('caught');
  end;
end."#
        ),
        &["cleanup", "caught"]
    );
}

#[test]
fn finally_assigns_flag_before_after_block() {
    assert_eq!(
        run_pascal(
            r#"program T;
var cleaned: Boolean;
begin
  cleaned := False;
  try
    WriteLn('work');
  finally
    cleaned := True;
  end;
  if cleaned then WriteLn('ok');
end."#
        ),
        &["work", "ok"]
    );
}

#[test]
fn nested_finally_inner_then_outer() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      WriteLn('inner-try');
    finally
      WriteLn('inner-finally');
    end;
  finally
    WriteLn('outer-finally');
  end;
end."#
        ),
        &["inner-try", "inner-finally", "outer-finally"]
    );
}

#[test]
fn finally_in_function_before_result_used() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Compute: Integer;
begin
  Result := 0;
  try
    Result := 6;
  finally
    WriteLn('fin');
  end;
end;
begin
  WriteLn(Compute);
end."#
        ),
        &["fin", "6"]
    );
}

#[test]
fn finally_runs_on_exit_from_try() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure StopEarly;
begin
  try
    WriteLn('start');
    Exit;
  finally
    WriteLn('always');
  end;
  WriteLn('never');
end;
begin
  StopEarly;
end."#
        ),
        &["start", "always"]
    );
}

#[test]
fn finally_with_except_outside_preserves_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('e1');
    finally
      WriteLn('f1');
    end;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["f1", "e1"]
    );
}

#[test]
fn finally_counter_increments_even_on_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 0;
  try
    try
      n := n + 1;
      raise Exception.Create('x');
    finally
      n := n + 1;
    end;
  except
    WriteLn(n);
  end;
end."#
        ),
        &["2"]
    );
}

// -------------------------------------------------------------------
// from test_exceptions_except_handlers.rs
// -------------------------------------------------------------------
#[test]
fn except_on_exception_writes_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('bad');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["bad"]
    );
}

#[test]
fn except_bare_handler_without_on_clause() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('x');
  except
    WriteLn('handled');
  end;
end."#
        ),
        &["handled"]
    );
}

#[test]
fn except_handler_runs_only_on_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('ok');
  except
    WriteLn('fail');
  end;
end."#
        ),
        &["ok"]
    );
}

#[test]
fn except_outer_catches_inner_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('inner');
    except
      on E: Exception do raise Exception.Create('outer:' + E.Message);
    end;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["outer:inner"]
    );
}

#[test]
fn except_in_loop_continues_after_handling() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  for i := 1 to 2 do begin
    try
      if i = 1 then raise Exception.Create('e');
      WriteLn('fine');
    except
      WriteLn('caught');
    end;
  end;
end."#
        ),
        &["caught", "fine"]
    );
}

#[test]
fn except_division_by_zero_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x, y, z: Integer;
begin
  x := 1; y := 0;
  try
    z := x div y;
    WriteLn(z);
  except
    on E: Exception do WriteLn('div0');
  end;
end."#
        ),
        &["div0"]
    );
}

#[test]
fn except_custom_exception_class_name() {
    assert_eq!(
        run_pascal(
            r#"program T;
type EMyError = class(Exception);
begin
  try
    raise EMyError.Create('mine');
  except
    on E: EMyError do WriteLn('my');
  end;
end."#
        ),
        &["my"]
    );
}

#[test]
fn except_code_after_try_except_still_runs() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('e');
  except
    WriteLn('h');
  end;
  WriteLn('after');
end."#
        ),
        &["h", "after"]
    );
}

#[test]
fn except_on_base_type_catches_derived_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
type ECustom = class(Exception);
begin
  try
    raise ECustom.Create('derived');
  except
    on E: Exception do WriteLn(E.ClassName);
  end;
end."#
        ),
        &["ECustom"]
    );
}

#[test]
fn except_finally_except_order_on_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('x');
    finally
      WriteLn('fin');
    end;
  except
    WriteLn('ex');
  end;
end."#
        ),
        &["fin", "ex"]
    );
}

#[test]
fn except_free_exception_object_in_handler() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('msg');
  except
    on E: Exception do begin
      WriteLn(E.Message);
      E.Free;
    end;
  end;
  WriteLn('ok');
end."#
        ),
        &["msg", "ok"]
    );
}

#[test]
fn except_raise_from_except_rewrapped() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
  try
    raise Exception.Create('a');
  except
    on E: Exception do raise Exception.Create('b:' + E.Message);
  end;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["b:a"]
    );
}

#[test]
fn except_try_body_assignment_before_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n: Integer;
begin
  n := 0;
  try
    n := 5;
    raise Exception.Create('stop');
  except
    WriteLn(n);
  end;
end."#
        ),
        &["5"]
    );
}

#[test]
fn except_handler_does_not_run_without_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('body');
  except
    WriteLn('handler');
  end;
end."#
        ),
        &["body"]
    );
}

#[test]
fn except_nested_finally_in_except_still_runs() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('e');
    except
      on E: Exception do begin
        try
          WriteLn('inner');
        finally
          WriteLn('cleanup');
        end;
      end;
    end;
  except
    WriteLn('outer');
  end;
end."#
        ),
        &["inner", "cleanup"]
    );
}

#[test]
fn except_procedure_raise_bubbles_to_caller() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure Boom;
begin
  raise Exception.Create('proc');
end;
begin
  try
    Boom;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["proc"]
    );
}

#[test]
fn except_finally_runs_after_except_handler() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('e');
    except
      WriteLn('handled');
    end;
  finally
    WriteLn('always');
  end;
end."#
        ),
        &["handled", "always"]
    );
}

#[test]
fn except_empty_except_clause_swallows_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('gone');
  except
  end;
  WriteLn('continued');
end."#
        ),
        &["continued"]
    );
}

#[test]
fn raise_out_of_memory_class_name() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise EOutOfMemory.Create('oom');
  except
    on E: EOutOfMemory do WriteLn(E.ClassName);
  end;
end."#
        ),
        &["EOutOfMemory"]
    );
}

#[test]
fn re_raise_propagates_to_outer_handler() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('x');
    except
      raise;
    end;
  except
    on E: Exception do WriteLn('outer');
  end;
end."#
        ),
        &["outer"]
    );
}

#[test]
fn except_on_base_class_catches_derived() {
    assert_eq!(
        run_pascal(
            r#"program T;
type EChild = class(Exception);
begin
  try
    raise EChild.Create('c');
  except
    on E: Exception do WriteLn('base-handler');
  end;
end."#
        ),
        &["base-handler"]
    );
}

#[test]
fn try_except_does_not_run_finally_before_handler() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('e');
    except
      WriteLn('handled');
    end;
  finally
    WriteLn('fin');
  end;
end."#
        ),
        &["handled", "fin"]
    );
}

#[test]
fn exception_message_preserved_after_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('msg123');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["msg123"]
    );
}


