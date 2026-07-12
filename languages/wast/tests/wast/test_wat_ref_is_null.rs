use crate::wat_exec;

wat_exec! {
    test_ref_is_null_true_funcref => { r#"
(func (export "_start")
  ref.null func
  ref.is_null
  call $log
)
"#, "1" },

    test_ref_is_null_true_externref => { r#"
(func (export "_start")
  ref.null extern
  ref.is_null
  call $log
)
"#, "1" },

    test_ref_is_null_false_funcref => { r#"
(func $f1)
(func (export "_start")
  ref.func $f1
  ref.is_null
  call $log
)
"#, "0" },

    test_ref_is_null_false_struct => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  ref.is_null
  call $log
)
"#, "0" },

    test_ref_is_null_true_struct => { r#"
(type $S (struct (field i32)))
(func (export "_start")
  ref.null $S
  ref.is_null
  call $log
)
"#, "1" },

    test_ref_is_null_after_local_set => { r#"
(type $S (struct (field i32)))
(func (export "_start") (local $s (ref null $S))
  ref.null $S
  local.set $s
  local.get $s
  ref.is_null
  call $log
)
"#, "1" },

    test_ref_is_null_after_local_set_non_null => { r#"
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

    test_ref_is_null_branch => { r#"
(func (export "_start")
  ref.null func
  ref.is_null
  if
    i32.const 42
    call $log
  else
    i32.const 99
    call $log
  end
)
"#, "42" }
}
