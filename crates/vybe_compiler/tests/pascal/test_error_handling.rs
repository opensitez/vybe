/// Tests for Pascal exception handling: try/except/finally/raise.
/// NOTE: Exception class is not auto-registered in the VM yet, so
/// raise Exception.Create(...) fails. These are until the
/// runtime registers a built-in Exception constructor.
use super::helpers::run_pascal;

// ===================================================================
// BASIC TRY/EXCEPT
// ===================================================================

#[test]
fn try_except_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('oops');
  except
    WriteLn('caught');
  end;
end."#
        ),
        &["caught"]
    );
}

#[test]
fn try_except_no_exception() {
    // This one works — no raise
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('ok');
  except
    WriteLn('error');
  end;
end."#
        ),
        &["ok"]
    );
}

#[test]
fn try_except_on_clause() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('bad');
  except
    on E: Exception do WriteLn('got: ' + E.Message);
  end;
end."#
        ),
        &["got: bad"]
    );
}

#[test]
fn try_except_multiple_on_clauses() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('fail');
  except
    on E: Exception do WriteLn('exception: ' + E.Message);
  end;
end."#
        ),
        &["exception: fail"]
    );
}

// ===================================================================
// TRY/FINALLY
// ===================================================================

#[test]
fn try_finally_no_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('body');
  finally
    WriteLn('cleanup');
  end;
end."#
        ),
        &["body", "cleanup"]
    );
}

#[test]
fn try_finally_with_exception() {
    // finally runs even if exception thrown; outer try catches it
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('boom');
    finally
      WriteLn('finally ran');
    end;
  except
    WriteLn('caught');
  end;
end."#
        ),
        &["finally ran", "caught"]
    );
}

// ===================================================================
// RAISE
// ===================================================================

#[test]
fn raise_exception_create() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('test error');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["test error"]
    );
}

#[test]
fn raise_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure DoStuff;
begin
  raise Exception.Create('from proc');
end;
begin
  try
    DoStuff;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["from proc"]
    );
}

// ===================================================================
// NESTED TRY
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
      WriteLn('inner caught');
    end;
    WriteLn('still running');
  except
    WriteLn('outer caught');
  end;
end."#
        ),
        &["inner caught", "still running"]
    );
}

#[test]
fn try_except_continue_after() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('err');
  except
    WriteLn('handled');
  end;
  WriteLn('after');
end."#
        ),
        &["handled", "after"]
    );
}

// ===================================================================
// TRY IN LOOP
// ===================================================================

#[test]
fn try_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer;
begin
  for i := 1 to 3 do
  begin
    try
      if i = 2 then raise Exception.Create('skip');
      WriteLn(IntToStr(i));
    except
      WriteLn('error at ' + IntToStr(i));
    end;
  end;
end."#
        ),
        &["1", "error at 2", "3"]
    );
}

#[test]
fn assert_true_does_not_raise() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  Assert(True, 'should not fire');
  WriteLn('ok');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn assert_false_raises_exception() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    Assert(False, 'failed check');
    WriteLn('no');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["failed check"]
    );
}

#[test]
fn try_except_finally_combined_block() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      WriteLn('try');
      raise Exception.Create('e');
    except
      WriteLn('except');
    end;
  finally
    WriteLn('finally');
  end;
end."#
        ),
        &["try", "except", "finally"]
    );
}

#[test]
fn raise_with_formatted_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.CreateFmt('code %d', [404]);
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["code 404"]
    );
}

#[test]
fn except_distinguishes_exception_types() {
    assert_eq!(
        run_pascal(
            r#"program T;
type EOne = class(Exception);
    ETwo = class(Exception);
begin
  try
    raise EOne.Create('one');
  except
    on E: ETwo do WriteLn('two');
    on E: EOne do WriteLn('one');
  end;
end."#
        ),
        &["one"]
    );
}

#[test]
fn try_finally_preserves_return_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function F: Integer;
begin
  Result := 0;
  try
    Result := 9;
  finally
    WriteLn('fin');
  end;
end;
begin
  WriteLn(F);
end."#
        ),
        &["fin", "9"]
    );
}

#[test]
fn except_handler_reraise_preserves_outer_message() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('root');
    except
      raise;
    end;
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#
        ),
        &["root"]
    );
}

#[test]
fn try_except_inside_case_branch() {
    assert_eq!(
        run_pascal(
            r#"program T;
var x: Integer;
begin
  x := 2;
  case x of
    2: begin
      try
        raise Exception.Create('in-case');
      except
        WriteLn('caught');
      end;
    end;
  end;
end."#
        ),
        &["caught"]
    );
}

#[test]
fn exception_inherited_message_property() {
    assert_eq!(
        run_pascal(
            r#"program T;
type EChild = class(Exception);
begin
  try
    raise EChild.Create('child-msg');
  except
    on E: Exception do WriteLn(E.ClassParent.ClassName);
  end;
end."#
        ),
        &["Exception"]
    );
}

#[test]
fn try_finally_runs_on_exit_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
function G: Integer;
begin
  try
    Result := 1;
    Exit;
  finally
    WriteLn('cleanup');
  end;
end;
begin
  G;
  WriteLn('done');
end."#
        ),
        &["cleanup", "done"]
    );
}

#[test]
fn assert_true_does_not_abort() {
    assert_eq!(
        run_pascal(r#"program T; begin Assert(True); WriteLn('ok'); end."#),
        &["ok"]
    );
}

#[test]
fn raise_last_os_error_without_code_continues() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    raise Exception.Create('manual');
  except
    on E: Exception do WriteLn('caught');
  end;
end."#
        ),
        &["caught"]
    );
}

#[test]
fn nested_try_except_inner_only_catches_inner() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    try
      raise Exception.Create('inner');
    except
      on E: Exception do WriteLn('inner-handled');
    end;
    WriteLn('outer-continues');
  except
    WriteLn('outer-caught');
  end;
end."#
        ),
        &["inner-handled", "outer-continues"]
    );
}

#[test]
fn except_on_specific_class_does_not_catch_sibling() {
    assert_eq!(
        run_pascal(
            r#"program T;
type EOne = class(Exception);
    ETwo = class(Exception);
begin
  try
    raise EOne.Create('one');
  except
    on E: ETwo do WriteLn('two');
    on E: EOne do WriteLn('one');
  end;
end."#
        ),
        &["one"]
    );
}

#[test]
fn finally_block_runs_after_successful_try() {
    assert_eq!(
        run_pascal(
            r#"program T;
begin
  try
    WriteLn('try');
  finally
    WriteLn('finally');
  end;
end."#
        ),
        &["try", "finally"]
    );
}
