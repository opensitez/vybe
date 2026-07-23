use crate::wat_exec;

wat_exec! {
    // A global initialized with `(ref.null $t)` holds a WASM GC typed null, so
    // struct.get on it traps per spec (not a lenient read).
    test_global_ref_null_traps => { r#"
(type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  global.get $g
  struct.get $S 0
  call $log
)
"#, "trap" },

    // ref.is_null on a ref.null-initialized global is 1.
    test_global_ref_null_is_null => { r#"
(type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  (ref.is_null (global.get $g))
  call $log
)
"#, "1" },

    test_global_local_mut => { r#"
(global $g (mut i32) (i32.const 10))
(func (export "_start")
  global.get $g
  i32.const 5
  i32.add
  global.set $g
  global.get $g
  call $log
)
"#, "15" },

    test_global_local_const => { r#"
(global $g i32 (i32.const 42))
(func (export "_start")
  global.get $g
  call $log
)
"#, "42" },

    test_global_multiple => { r#"
(global $g1 (mut i32) (i32.const 10))
(global $g2 (mut i32) (i32.const 20))
(func (export "_start")
  global.get $g1
  global.get $g2
  i32.add
  call $log
)
"#, "30" },

    test_global_init_expr => { r#"
(global $g1 i32 (i32.const 10))
(global $g2 i32 (global.get $g1))
(func (export "_start")
  global.get $g2
  call $log
)
"#, "10" },

    test_global_different_types => { r#"
(global $gi i32 (i32.const 42))
(global $gf f32 (f32.const 3.14))
(global $gl i64 (i64.const 99))
(func (export "_start")
  global.get $gl
  call $log_i64
)
"#, "99" },

    test_global_export => { r#"
(global $g (export "g") (mut i32) (i32.const 42))
(func (export "_start")
  global.get $g
  call $log
)
"#, "42" },

    test_global_ref_type => { r#"
(global $g (mut funcref) (ref.null func))
(func $f1)
(func (export "_start")
  ref.func $f1
  global.set $g
  global.get $g
  ref.is_null
  call $log
)
"#, "0" },

    test_global_struct_type => { r#"
(type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  i32.const 42
  struct.new $S
  global.set $g
  global.get $g
  struct.get $S 0
  call $log
)
"#, "42" }
}
