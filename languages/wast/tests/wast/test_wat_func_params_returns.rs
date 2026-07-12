use crate::wat_exec;

wat_exec! {
    test_param_get => { r#"
(func $f1 (param $x i32) (result i32)
  local.get $x)
(func (export "_start")
  i32.const 42
  call $f1
  call $log
)
"#, "42" },

    test_param_multiple => { r#"
(func $f1 (param $x i32) (param $y i32) (result i32)
  local.get $x
  local.get $y
  i32.add)
(func (export "_start")
  i32.const 10
  i32.const 20
  call $f1
  call $log
)
"#, "30" },

    test_param_mutate => { r#"
(func $f1 (param $x i32) (result i32)
  local.get $x
  i32.const 10
  i32.add
  local.set $x
  local.get $x)
(func (export "_start")
  i32.const 42
  call $f1
  call $log
)
"#, "52" },

    test_param_tee => { r#"
(func $f1 (param $x i32) (result i32)
  local.get $x
  i32.const 10
  i32.add
  local.tee $x)
(func (export "_start")
  i32.const 42
  call $f1
  call $log
)
"#, "52" },

    test_param_types => { r#"
(func $f1 (param $i i32) (param $f f32) (param $l i64) (param $d f64) (result f64)
  local.get $d)
(func (export "_start")
  i32.const 10
  f32.const 1.0
  i64.const 20
  f64.const 42.5
  call $f1
  call $log_f64
)
"#, "42.5" },

    test_return_void => { r#"
(func $f1
  nop)
(func (export "_start")
  call $f1
  i32.const 42
  call $log
)
"#, "42" },

    test_return_early => { r#"
(func $f1 (result i32)
  i32.const 42
  return
  i32.const 99)
(func (export "_start")
  call $f1
  call $log
)
"#, "42" },

    test_return_early_nested => { r#"
(func $f1 (result i32)
  block
    block
      i32.const 42
      return
    end
  end
  i32.const 99)
(func (export "_start")
  call $f1
  call $log
)
"#, "42" },

    test_return_multiple => { r#"
(func $f1 (result i32 i32)
  i32.const 10
  i32.const 20)
(func (export "_start")
  call $f1
  i32.add
  call $log
)
"#, "30" },

    test_return_multiple_early => { r#"
(func $f1 (result i32 i32)
  i32.const 10
  i32.const 20
  return
  i32.const 99
  i32.const 99)
(func (export "_start")
  call $f1
  i32.add
  call $log
)
"#, "30" },

    test_return_empty => { r#"
(func $f1
  return)
(func (export "_start")
  call $f1
  i32.const 42
  call $log
)
"#, "42" },

    test_param_index_access => { r#"
(func $f1 (param i32) (param i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start")
  i32.const 10
  i32.const 20
  call $f1
  call $log
)
"#, "30" },

    test_param_ref_null => { r#"
(type $S (struct (field i32)))
(func $f1 (param $s (ref null $S)) (result i32)
  local.get $s
  ref.is_null)
(func (export "_start")
  ref.null $S
  call $f1
  call $log
)
"#, "1" },

    test_param_ref_non_null => { r#"
(type $S (struct (field i32)))
(func $f1 (param $s (ref null $S)) (result i32)
  local.get $s
  ref.is_null)
(func (export "_start")
  i32.const 42
  struct.new $S
  call $f1
  call $log
)
"#, "0" }
}
