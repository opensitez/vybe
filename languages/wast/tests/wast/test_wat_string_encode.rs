use crate::wat_exec;

wat_exec! {
    test_string_encode_utf8 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 10
  string.encode_utf8
  drop
  
  i32.const 10
  i32.load8_u
  call $log
)
"#, "104" }, // 'h'

    test_string_encode_wtf16 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 10
  string.encode_wtf16
  drop
  
  i32.const 10
  i32.load16_u
  call $log
)
"#, "104" }, // 'h'

    test_string_encode_utf8_array => { r#"
(memory 1)
(type $A (array (mut i8)))
(data (i32.const 0) "hello")
(func (export "_start") (local $a (ref null $A))
  i32.const 5
  array.new_default $A
  local.set $a
  
  i32.const 0
  i32.const 5
  string.new_utf8
  
  local.get $a
  i32.const 0
  string.encode_utf8_array
  drop
  
  local.get $a
  i32.const 1
  array.get_u $A
  call $log
)
"#, "101" }, // 'e'

    test_string_encode_wtf16_array => { r#"
(memory 1)
(type $A (array (mut i16)))
(data (i32.const 0) "hello")
(func (export "_start") (local $a (ref null $A))
  i32.const 5
  array.new_default $A
  local.set $a
  
  i32.const 0
  i32.const 5
  string.new_utf8
  
  local.get $a
  i32.const 0
  string.encode_wtf16_array
  drop
  
  local.get $a
  i32.const 1
  array.get_u $A
  call $log
)
"#, "101" }, // 'e'

    test_string_encode_utf8_oob => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 65535
  string.encode_utf8
  drop
  
  i32.const 42
  call $log
)
"#, "trap" },

    test_string_encode_null => { r#"
(memory 1)
(func (export "_start")
  ref.null string
  i32.const 10
  string.encode_utf8
  drop
  i32.const 42
  call $log
)
"#, "trap" }
}
