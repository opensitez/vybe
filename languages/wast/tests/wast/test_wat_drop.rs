use crate::wat_exec;

wat_exec! {
    test_drop_i32 => { r#"
(func (export "_start")
  i32.const 10
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_i64 => { r#"
(func (export "_start")
  i64.const 10
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_f32 => { r#"
(func (export "_start")
  f32.const 10.0
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_f64 => { r#"
(func (export "_start")
  f64.const 10.0
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_ref => { r#"
(func $f1)
(func (export "_start")
  ref.func $f1
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_multiple => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  drop
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_in_block => { r#"
(func (export "_start")
  block
    i32.const 10
    drop
  end
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_in_if => { r#"
(func (export "_start")
  i32.const 1
  if
    i32.const 10
    drop
  end
  i32.const 42
  call $log
)
"#, "42" },
    
    test_drop_function_result => { r#"
(func $f1 (result i32) i32.const 99)
(func (export "_start")
  call $f1
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_drop_block_result => { r#"
(func (export "_start")
  block (result i32)
    i32.const 99
  end
  drop
  i32.const 42
  call $log
)
"#, "42" }
}
