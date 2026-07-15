use crate::wat_exec;

wat_exec! {
    test_exception_throw_catch => { r#"
(tag $e (param i32))
(func (export "_start")
  try (result i32)
    i32.const 42
    throw $e
  catch $e
    ;; the payload is left on the stack
  end
  call $log
)
"#, "42" },

    test_exception_no_throw => { r#"
(tag $e (param i32))
(func (export "_start")
  try (result i32)
    i32.const 99
  catch $e
  end
  call $log
)
"#, "99" },

    test_exception_catch_all => { r#"
(tag $e (param i32))
(func (export "_start")
  try (result i32)
    i32.const 42
    throw $e
  catch_all
    i32.const 99
  end
  call $log
)
"#, "99" },

    test_exception_nested_try => { r#"
(tag $e1 (param i32))
(tag $e2 (param i32))
(func (export "_start")
  try (result i32)
    try (result i32)
      i32.const 42
      throw $e2
    catch $e1
      i32.const 10
    end
  catch $e2
    ;; caught by outer
  end
  call $log
)
"#, "42" },

    test_exception_delegate => { r#"
(tag $e (param i32))
(func (export "_start")
  try (result i32)
    try (result i32)
      i32.const 42
      throw $e
    delegate 0
  catch $e
  end
  call $log
)
"#, "42" },

    test_exception_throw_multiple_args => { r#"
(tag $e (param i32 i32))
(func (export "_start")
  try (result i32)
    i32.const 10
    i32.const 20
    throw $e
  catch $e
    i32.add
  end
  call $log
)
"#, "30" },

    test_exception_rethrow => { r#"
(tag $e (param i32))
(func (export "_start")
  try (result i32)
    try (result i32)
      i32.const 42
      throw $e
    catch $e
      rethrow 1 ;; rethrow the exception caught by this catch
    end
  catch $e
  end
  call $log
)
"#, "42" },

    // A `throw` raised inside a *called* function unwinds across the call
    // frame and is caught by the caller's `try`. The tag `$e` is a single
    // load-time entity shared by both functions' chunks, so the throw and the
    // catch resolve to the same tag — the unified `throw`/`try_table`
    // mechanism is compatible across compilation units.
    test_exception_throw_propagates_across_call => { r#"
(tag $e (param i32))
(func $raise
  i32.const 7
  throw $e
)
(func (export "_start")
  try (result i32)
    call $raise
    i32.const 0 ;; unreached
  catch $e
  end
  call $log
)
"#, "7" }
}
