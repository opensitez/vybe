use crate::wat_exec;

wat_exec! {
    test_call_simple => { r#"
(func $helper (result i32)
  i32.const 42
)
(func (export "_start")
  call $helper
  call $log
)
"#, "42" },

    test_call_with_args => { r#"
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add
)
(func (export "_start")
  i32.const 10
  i32.const 20
  call $add
  call $log
)
"#, "30" },

    test_call_recursive => { r#"
(func $fact (param $n i32) (result i32)
  (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
    (then (i32.const 1))
    (else
      (i32.mul
        (local.get $n)
        (call $fact (i32.sub (local.get $n) (i32.const 1)))
      )
    )
  )
)
(func (export "_start")
  i32.const 5
  call $fact
  call $log
)
"#, "120" }
}
