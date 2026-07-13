use crate::wat_exec;

wat_exec! {
    test_string_new_utf8 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  string.measure_utf8
  call $log
)
"#, "5" },

    test_string_new_utf8_empty => { r#"
(memory 1)
(func (export "_start")
  i32.const 0
  i32.const 0
  string.new_utf8
  string.measure_utf8
  call $log
)
"#, "0" },

    test_string_new_utf8_oob => { r#"
(memory 1)
(func (export "_start")
  i32.const 65530
  i32.const 10
  string.new_utf8
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_string_new_wtf8 => { r#"
(memory 1)
(data (i32.const 0) "world")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_wtf8
  string.measure_wtf8
  call $log
)
"#, "5" },

    test_string_new_utf8_array => { r#"
(type $A (array (mut i8)))
(func (export "_start") (local $a (ref null $A))
  i32.const 104 ;; 'h'
  i32.const 105 ;; 'i'
  array.new_fixed $A 2
  local.set $a
  
  local.get $a
  i32.const 0
  i32.const 2
  string.new_utf8_array
  string.measure_utf8
  call $log
)
"#, "2" },

    test_string_new_wtf16_array => { r#"
(type $A (array (mut i16)))
(func (export "_start") (local $a (ref null $A))
  i32.const 104 ;; 'h'
  i32.const 105 ;; 'i'
  array.new_fixed $A 2
  local.set $a
  
  local.get $a
  i32.const 0
  i32.const 2
  string.new_wtf16_array
  string.measure_utf8
  call $log
)
"#, "2" },
    
    test_string_new_lossy_utf8 => { r#"
(memory 1)
(data (i32.const 0) "\ff\ff") ;; invalid utf8
(func (export "_start")
  i32.const 0
  i32.const 2
  string.new_lossy_utf8
  string.measure_utf8
  call $log
)
"#, "6" } // replacement character \ufffd is 3 bytes in utf8, two of them = 6 bytes
}
