use crate::wat_exec;

wat_exec! {
    test_string_concat => { r#"
(memory 1)
(data (i32.const 0) "hello")
(data (i32.const 10) "world")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 10
  i32.const 5
  string.new_utf8
  
  string.concat
  string.measure_utf8
  call $log
)
"#, "10" },

    test_string_concat_empty_first => { r#"
(memory 1)
(data (i32.const 10) "world")
(func (export "_start")
  i32.const 0
  i32.const 0
  string.new_utf8
  
  i32.const 10
  i32.const 5
  string.new_utf8
  
  string.concat
  string.measure_utf8
  call $log
)
"#, "5" },

    test_string_concat_empty_second => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 10
  i32.const 0
  string.new_utf8
  
  string.concat
  string.measure_utf8
  call $log
)
"#, "5" },

    test_string_concat_empty_both => { r#"
(memory 1)
(func (export "_start")
  i32.const 0
  i32.const 0
  string.new_utf8
  
  i32.const 10
  i32.const 0
  string.new_utf8
  
  string.concat
  string.measure_utf8
  call $log
)
"#, "0" },

    test_string_concat_null_first => { r#"
(memory 1)
(data (i32.const 10) "world")
(func (export "_start")
  ref.null string
  
  i32.const 10
  i32.const 5
  string.new_utf8
  
  string.concat
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_string_concat_null_second => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  ref.null string
  
  string.concat
  drop
  i32.const 42
  call $log
)
"#, "trap" }
}
