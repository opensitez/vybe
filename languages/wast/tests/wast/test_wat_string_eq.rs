use crate::wat_exec;

wat_exec! {
    test_string_eq_same => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 0
  i32.const 5
  string.new_utf8
  
  string.eq
  call $log
)
"#, "1" },

    test_string_eq_diff_length => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  i32.const 0
  i32.const 4
  string.new_utf8
  
  string.eq
  call $log
)
"#, "0" },

    test_string_eq_diff_content => { r#"
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
  
  string.eq
  call $log
)
"#, "0" },

    test_string_eq_empty => { r#"
(memory 1)
(func (export "_start")
  i32.const 0
  i32.const 0
  string.new_utf8
  
  i32.const 10
  i32.const 0
  string.new_utf8
  
  string.eq
  call $log
)
"#, "1" },

    test_string_eq_null_first => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  ref.null string
  
  i32.const 0
  i32.const 5
  string.new_utf8
  
  string.eq
  call $log
)
"#, "0" },

    test_string_eq_null_second => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  ref.null string
  
  string.eq
  call $log
)
"#, "0" },

    test_string_eq_null_both => { r#"
(func (export "_start")
  ref.null string
  ref.null string
  string.eq
  call $log
)
"#, "1" },

    test_string_compare_lt => { r#"
(memory 1)
(data (i32.const 0) "abc")
(data (i32.const 10) "xyz")
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  
  i32.const 10
  i32.const 3
  string.new_utf8
  
  string.compare
  call $log
)
"#, "-1" },

    test_string_compare_gt => { r#"
(memory 1)
(data (i32.const 0) "xyz")
(data (i32.const 10) "abc")
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  
  i32.const 10
  i32.const 3
  string.new_utf8
  
  string.compare
  call $log
)
"#, "1" },

    test_string_compare_eq => { r#"
(memory 1)
(data (i32.const 0) "abc")
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  
  i32.const 0
  i32.const 3
  string.new_utf8
  
  string.compare
  call $log
)
"#, "0" },

    test_string_compare_diff_len => { r#"
(memory 1)
(data (i32.const 0) "abc")
(data (i32.const 10) "abcd")
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  
  i32.const 10
  i32.const 4
  string.new_utf8
  
  string.compare
  call $log
)
"#, "-1" }
}
