use crate::wat_exec;

wat_exec! {
    test_local_get_set => { r#"
(func (export "_start") (local $x i32)
  i32.const 42
  local.set $x
  local.get $x
  call $log
)
"#, "42" },

    test_local_tee => { r#"
(func (export "_start") (local $x i32)
  i32.const 42
  local.tee $x
  call $log
)
"#, "42" },

    test_local_tee_and_get => { r#"
(func (export "_start") (local $x i32)
  i32.const 42
  local.tee $x
  local.get $x
  i32.add
  call $log
)
"#, "84" },

    test_local_multiple => { r#"
(func (export "_start") (local $x i32) (local $y i32)
  i32.const 10
  local.set $x
  i32.const 20
  local.set $y
  local.get $x
  local.get $y
  i32.add
  call $log
)
"#, "30" },

    test_local_default_zero => { r#"
(func (export "_start") (local $x i32)
  local.get $x
  call $log
)
"#, "0" },

    test_local_f32_default_zero => { r#"
(func (export "_start") (local $x f32)
  local.get $x
  call $log_f32
)
"#, "0.0" },

    test_local_f64_default_zero => { r#"
(func (export "_start") (local $x f64)
  local.get $x
  call $log_f64
)
"#, "0.0" },

    test_local_i64_default_zero => { r#"
(func (export "_start") (local $x i64)
  local.get $x
  call $log_i64
)
"#, "0" },

    test_local_shadowing_param => { r#"
(func $f1 (param $x i32) (result i32) (local $y i32)
  local.get $x
  local.set $y
  local.get $y)
(func (export "_start")
  i32.const 42
  call $f1
  call $log
)
"#, "42" },
    
    test_local_shadowing_global => { r#"
(global $g (mut i32) (i32.const 10))
(func (export "_start") (local $g i32)
  i32.const 42
  local.set $g
  local.get $g
  call $log
)
"#, "42" },

    test_local_ref_null => { r#"
(func (export "_start") (local $r funcref)
  local.get $r
  ref.is_null
  call $log
)
"#, "1" },

    test_local_struct_null => { r#"
(type $S (struct (field i32)))
(func (export "_start") (local $s (ref null $S))
  local.get $s
  ref.is_null
  call $log
)
"#, "1" },

    test_local_struct_set => { r#"
(type $S (struct (field i32)))
(func (export "_start") (local $s (ref null $S))
  i32.const 42
  struct.new $S
  local.set $s
  local.get $s
  ref.is_null
  call $log
)
"#, "0" },

    test_local_array_null => { r#"
(type $A (array i32))
(func (export "_start") (local $a (ref null $A))
  local.get $a
  ref.is_null
  call $log
)
"#, "1" },
    
    test_local_index_access => { r#"
(func (export "_start") (local i32) (local i32)
  i32.const 10
  local.set 0
  i32.const 20
  local.set 1
  local.get 0
  local.get 1
  i32.add
  call $log
)
"#, "30" }
}
