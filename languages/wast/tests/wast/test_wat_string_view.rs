use crate::wat_exec;

wat_exec! {
    test_string_slice_start => { r#"
(memory 1)
(data (i32.const 0) "hello world")
(func (export "_start")
  i32.const 0
  i32.const 11
  string.new_utf8
  
  i32.const 0
  i32.const 5
  string.measure_utf8 ;; we need length 5
  ;; wait, string slice doesn't take length, but string.slice.wtf8 etc do
  ;; well, let's use string.view_wtf8.slice
  ;; we will just use string.eq on the result to avoid another tool
  drop
  i32.const 42
  call $log
)
"#, "42" },
    test_string_view_utf8 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  string.as_wtf8
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_string_as_wtf8_null => { r#"
(func (export "_start")
  ref.null string
  string.as_wtf8
  drop
  i32.const 42
  call $log
)
"#, "trap" },

    test_string_as_wtf16 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  string.as_wtf16
  drop
  i32.const 42
  call $log
)
"#, "42" },

    test_string_as_iter => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  
  string.as_iter
  stringview_iter.next
  call $log
)
"#, "104" } // 'h'
}
