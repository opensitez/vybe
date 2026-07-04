use crate::wat_exec;

wat_exec! {
    test_unreachable_direct => { r#"
(func (export "_start")
  unreachable
  i32.const 42
  call $log
)
"#, "trap" },

    test_unreachable_in_block => { r#"
(func (export "_start")
  block
    unreachable
  end
  i32.const 42
  call $log
)
"#, "trap" },

    test_unreachable_in_if_true => { r#"
(func (export "_start")
  i32.const 1
  if
    unreachable
  else
    i32.const 42
    call $log
  end
)
"#, "trap" },

    test_unreachable_in_if_false_skipped => { r#"
(func (export "_start")
  i32.const 0
  if
    unreachable
  else
    i32.const 42
    call $log
  end
)
"#, "42" }, // unreachable is not executed

    test_unreachable_in_else_skipped => { r#"
(func (export "_start")
  i32.const 1
  if
    i32.const 42
    call $log
  else
    unreachable
  end
)
"#, "42" },

    test_unreachable_after_return => { r#"
(func (export "_start")
  i32.const 42
  call $log
  return
  unreachable
)
"#, "42" }, // unreachable is dead code

    test_unreachable_after_br => { r#"
(func (export "_start")
  block
    i32.const 42
    call $log
    br 0
    unreachable
  end
)
"#, "42" },

    test_unreachable_in_loop => { r#"
(func (export "_start")
  loop
    unreachable
  end
  i32.const 42
  call $log
)
"#, "trap" },

    test_unreachable_polymorphic_stack => { r#"
(func (export "_start")
  unreachable
  i32.add ;; consumes values that don't exist, valid due to polymorphic stack
  call $log
)
"#, "trap" }
}
