use crate::wat_exec;

wat_exec! {
    test_struct_set => { r#"
(type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  
  local.get $p
  i32.const 42
  struct.set $Point 0
  
  local.get $p
  struct.get $Point 0
  call $log
)
"#, "42" },

    test_struct_set_second => { r#"
(type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  
  local.get $p
  i32.const 99
  struct.set $Point 1
  
  local.get $p
  struct.get $Point 1
  call $log
)
"#, "99" },

    test_struct_set_null => { r#"
(type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  ref.null $Point
  local.set $p
  
  local.get $p
  i32.const 42
  struct.set $Point 0
  
  i32.const 0
  call $log
)
"#, "trap" },

    test_struct_set_multiple => { r#"
(type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  
  local.get $p
  i32.const 42
  struct.set $Point 0
  
  local.get $p
  i32.const 99
  struct.set $Point 1
  
  local.get $p
  struct.get $Point 0
  local.get $p
  struct.get $Point 1
  i32.add
  call $log
)
"#, "141" }
}
