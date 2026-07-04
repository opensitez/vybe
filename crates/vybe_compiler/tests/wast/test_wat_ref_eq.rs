use crate::wat_exec;

wat_exec! {
    test_ref_eq_null_null => { r#"
(func (export "_start")
  ref.null func
  ref.null func
  ref.eq
  call $log
)
"#, "1" },

    test_ref_eq_null_null_extern => { r#"
(func (export "_start")
  ref.null extern
  ref.null extern
  ref.eq
  call $log
)
"#, "1" },

    test_ref_eq_same_func => { r#"
(func $f1)
(func (export "_start")
  ref.func $f1
  ref.func $f1
  ref.eq
  call $log
)
"#, "1" },

    test_ref_eq_diff_func => { r#"
(func $f1)
(func $f2)
(func (export "_start")
  ref.func $f1
  ref.func $f2
  ref.eq
  call $log
)
"#, "0" },

    test_ref_eq_same_struct => { r#"
(type $S (struct (field i32)))
(func (export "_start") (local $s (ref null $S))
  i32.const 42
  struct.new $S
  local.tee $s
  local.get $s
  ref.eq
  call $log
)
"#, "1" },

    test_ref_eq_diff_struct_same_value => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  i32.const 42
  struct.new $S
  ref.eq
  call $log
)
"#, "0" }, // Different allocations

    test_ref_eq_struct_null => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  ref.null $S
  ref.eq
  call $log
)
"#, "0" },

    test_ref_eq_same_array => { r#"
(type $A (array i32))
(func (export "_start") (local $a (ref null $A))
  i32.const 42
  i32.const 5
  array.new $A
  local.tee $a
  local.get $a
  ref.eq
  call $log
)
"#, "1" },

    test_ref_eq_diff_array_same_value => { r#"
(type $A (array i32))
(func (export "_start")
  i32.const 42
  i32.const 5
  array.new $A
  i32.const 42
  i32.const 5
  array.new $A
  ref.eq
  call $log
)
"#, "0" } // Different allocations
}
