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
