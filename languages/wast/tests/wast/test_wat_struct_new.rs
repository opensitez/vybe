use crate::wat_exec;

wat_exec! {
    test_struct_new_default => { r#"
(type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  struct.new_default $Point
  local.set $p
  local.get $p
  struct.get $Point 0
  call $log
)
"#, "0" },

    test_struct_new_with_args => { r#"
(type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  local.get $p
  struct.get $Point 1
  call $log
)
"#, "20" },

    test_struct_new_nested => { r#"
(type $Point (struct (field i32) (field i32)))
(type $Rect (struct (field (ref $Point)) (field (ref $Point))))
(func (export "_start") (local $r (ref null $Rect))
  i32.const 10
  i32.const 20
  struct.new $Point
  
  i32.const 30
  i32.const 40
  struct.new $Point
  
  struct.new $Rect
  local.set $r
  
  local.get $r
  struct.get $Rect 1
  struct.get $Point 0
  call $log
)
"#, "30" },

    test_struct_new_mixed_types => { r#"
(type $Mixed (struct (field i32) (field f32) (field i64) (field f64)))
(func (export "_start") (local $m (ref null $Mixed))
  i32.const 42
  f32.const 3.14
  i64.const 99
  f64.const 2.71
  struct.new $Mixed
  local.set $m
  
  local.get $m
  struct.get $Mixed 2
  call $log_i64
)
"#, "99" }
}
