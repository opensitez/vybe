use crate::wat_exec;

wat_exec! {
    test_struct_subtype_get => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  local.get $s
  struct.get $Base 0
  call $log
)
"#, "10" },

    test_struct_subtype_mut => { r#"
(type $Base (struct (field (mut i32))))
(type $Sub (struct_subtype (field (mut i32)) (field (mut i32)) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  local.get $s
  i32.const 42
  struct.set $Base 0
  
  local.get $s
  struct.get $Base 0
  call $log
)
"#, "42" },

    test_struct_subtype_deep => { r#"
(type $Base (struct (field i32)))
(type $Sub1 (struct_subtype (field i32) (field i32) $Base))
(type $Sub2 (struct_subtype (field i32) (field i32) (field i32) $Sub1))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  i32.const 30
  struct.new $Sub2
  local.set $s
  
  local.get $s
  struct.get $Base 0
  call $log
)
"#, "10" },

    test_struct_subtype_func_param => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func $f1 (param $b (ref null $Base)) (result i32)
  local.get $b
  struct.get $Base 0)
(func (export "_start")
  i32.const 99
  i32.const 88
  struct.new $Sub
  call $f1
  call $log
)
"#, "99" },

    test_struct_subtype_func_return => { r#"
(type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func $f1 (result (ref null $Base))
  i32.const 42
  i32.const 88
  struct.new $Sub)
(func (export "_start")
  call $f1
  struct.get $Base 0
  call $log
)
"#, "42" }
}
