use crate::wat_exec;

wat_exec! {
    test_ref_as_non_null_struct => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  ref.as_non_null
  struct.get $S 0
  call $log
)
"#, "42" },

    test_ref_as_non_null_trap_struct => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  ref.null $S
  ref.as_non_null
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_ref_as_non_null_func => { r#"
(func $f1)
(func (export "_start")
  ref.func $f1
  ref.as_non_null
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_ref_as_non_null_trap_func => { r#"
(func (export "_start")
  ref.null func
  ref.as_non_null
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_ref_as_non_null_extern => { r#"
(func (export "_start")
  ref.null extern
  ref.as_non_null
  drop
  i32.const 42
  call $log
)
"#, "trap" }
}
