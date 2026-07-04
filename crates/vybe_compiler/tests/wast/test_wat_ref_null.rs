use crate::wat_exec;

wat_exec! {
    test_ref_is_null_true => { r#"
(type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (ref.null $t))
  (ref.is_null (local.get $s))
  call $log
)
"#, "1" },

    test_ref_is_null_false => { r#"
(type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (struct.new $t (i32.const 42)))
  (ref.is_null (local.get $s))
  call $log
)
"#, "0" },

    test_br_on_null_success => { r#"
(type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (ref.null $t))
  (block $L
    (drop (br_on_null $L (local.get $s)))
    (i32.const 0)
    call $log
    return
  )
  (i32.const 1)
  call $log
)
"#, "1" },

    test_br_on_non_null_success => { r#"
(type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (struct.new $t (i32.const 42)))
  (block $L (result (ref $t))
    (br_on_non_null $L (local.get $s))
    (return (i32.const 0))
  )
  (struct.get $t 0)
  call $log
)
"#, "42" }
}
