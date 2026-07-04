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
"#, "42" }
}
