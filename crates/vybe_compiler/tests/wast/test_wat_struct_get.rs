use crate::wat_exec;

wat_exec! {
    test_struct_get_first => { r#"
(type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  i32.const 10
  i32.const 20
  struct.new $Point
  local.set $p
  
  local.get $p
  struct.get $Point 0
  call $log
)
"#, "10" },

    test_struct_get_second => { r#"
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

    test_struct_get_null => { r#"
(type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  ref.null $Point
  local.set $p
  
  local.get $p
  struct.get $Point 0
  call $log
)
"#, "trap" },

    test_struct_get_s => { r#"
(type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; -1 as i8
  i32.const 65535 ;; -1 as i16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_s $S 0
  call $log
)
"#, "-1" },

    test_struct_get_u => { r#"
(type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; 255 as u8
  i32.const 65535 ;; 65535 as u16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_u $S 0
  call $log
)
"#, "255" },
    
    test_struct_get_s_16 => { r#"
(type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; -1 as i8
  i32.const 65535 ;; -1 as i16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_s $S 1
  call $log
)
"#, "-1" },

    test_struct_get_u_16 => { r#"
(type $S (struct (field i8) (field i16)))
(func (export "_start") (local $s (ref null $S))
  i32.const 255 ;; 255 as u8
  i32.const 65535 ;; 65535 as u16
  struct.new $S
  local.set $s
  
  local.get $s
  struct.get_u $S 1
  call $log
)
"#, "65535" }
}
