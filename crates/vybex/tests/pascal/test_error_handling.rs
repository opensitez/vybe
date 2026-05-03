/// Tests for Pascal exception handling: try/except/finally/raise.
/// NOTE: Exception class is not auto-registered in the VM yet, so
/// raise Exception.Create(...) fails. These are until the
/// runtime registers a built-in Exception constructor.

use super::helpers::run_pascal;

// ===================================================================
// BASIC TRY/EXCEPT
// ===================================================================

#[test] fn try_except_basic() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    raise Exception.Create('oops');
  except
    WriteLn('caught');
  end;
end."#), &["caught"]);
}

#[test] fn try_except_no_exception() { // This one works — no raise
    assert_eq!(run_pascal(r#"program T;
begin
  try
    WriteLn('ok');
  except
    WriteLn('error');
  end;
end."#), &["ok"]);
}

#[test] fn try_except_on_clause() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    raise Exception.Create('bad');
  except
    on E: Exception do WriteLn('got: ' + E.Message);
  end;
end."#), &["got: bad"]);
}

#[test] fn try_except_multiple_on_clauses() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    raise Exception.Create('fail');
  except
    on E: Exception do WriteLn('exception: ' + E.Message);
  end;
end."#), &["exception: fail"]);
}

// ===================================================================
// TRY/FINALLY
// ===================================================================

#[test] fn try_finally_no_exception() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    WriteLn('body');
  finally
    WriteLn('cleanup');
  end;
end."#), &["body", "cleanup"]);
}

#[test] fn try_finally_with_exception() {
    // finally runs even if exception thrown; outer try catches it
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["finally ran", "caught"]);
}

// ===================================================================
// RAISE
// ===================================================================

#[test] fn raise_exception_create() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    raise Exception.Create('test error');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end."#), &["test error"]);
}

#[test] fn raise_in_function() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["from proc"]);
}

// ===================================================================
// NESTED TRY
// ===================================================================

#[test] fn nested_try_except() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["inner caught", "still running"]);
}

#[test] fn try_except_continue_after() {
    assert_eq!(run_pascal(r#"program T;
begin
  try
    raise Exception.Create('err');
  except
    WriteLn('handled');
  end;
  WriteLn('after');
end."#), &["handled", "after"]);
}

// ===================================================================
// TRY IN LOOP
// ===================================================================

#[test] fn try_in_loop() {
    assert_eq!(run_pascal(r#"program T;
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
end."#), &["1", "error at 2", "3"]);
}
