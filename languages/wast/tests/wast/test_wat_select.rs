use crate::wat_exec;

wat_exec! {
    test_select_true => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const 1
  select
  call $log
)
"#, "10" },

    test_select_false => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const 0
  select
  call $log
)
"#, "20" },

    test_select_negative_condition => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const -1
  select
  call $log
)
"#, "10" }, // Any non-zero is true

    test_select_f32_true => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 2.71
  i32.const 1
  select
  call $log_f32
)
"#, "3.14" },

    test_select_f32_false => { r#"
(func (export "_start")
  f32.const 3.14
  f32.const 2.71
  i32.const 0
  select
  call $log_f32
)
"#, "2.71" },

    test_select_f64_true => { r#"
(func (export "_start")
  f64.const 3.14
  f64.const 2.71
  i32.const 42
  select
  call $log_f64
)
"#, "3.14" },

    test_select_i64_false => { r#"
(func (export "_start")
  i64.const 99
  i64.const 42
  i32.const 0
  select
  call $log_i64
)
"#, "42" },

    test_select_with_type_true => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const 1
  select (result i32)
  call $log
)
"#, "10" },

    test_select_with_type_false => { r#"
(func (export "_start")
  i32.const 10
  i32.const 20
  i32.const 0
  select (result i32)
  call $log
)
"#, "20" },

    test_select_ref_true => { r#"
(func $f1)
(func $f2)
(func (export "_start")
  ref.func $f1
  ref.func $f2
  i32.const 1
  select (result funcref)
  ref.is_null
  call $log
)
"#, "0" },

    test_select_ref_null => { r#"
(func $f1)
(func (export "_start")
  ref.func $f1
  ref.null func
  i32.const 0
  select (result funcref)
  ref.is_null
  call $log
)
"#, "1" }
}
