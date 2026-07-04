use crate::wat_exec;

wat_exec! {
    test_string_measure_utf8 => { r#"
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

    test_string_measure_wtf8 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_wtf8
  string.measure_wtf8
  call $log
)
"#, "5" },

    test_string_measure_wtf16 => { r#"
(memory 1)
(data (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.const 5
  string.new_utf8
  string.measure_wtf16
  call $log
)
"#, "5" }, ;; ascii characters take 1 code unit in wtf16

    test_string_measure_utf8_multibyte => { r#"
(memory 1)
(data (i32.const 0) "\e2\82\ac") ;; euro sign, 3 bytes
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  string.measure_utf8
  call $log
)
"#, "3" },

    test_string_measure_wtf16_multibyte => { r#"
(memory 1)
(data (i32.const 0) "\e2\82\ac") ;; euro sign, 3 bytes in utf8, 1 code unit in wtf16
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  string.measure_wtf16
  call $log
)
"#, "1" },

    test_string_measure_utf8_null => { r#"
(func (export "_start")
  ref.null string
  string.measure_utf8
  call $log
)
"#, "trap" },

    test_string_measure_wtf16_null => { r#"
(func (export "_start")
  ref.null string
  string.measure_wtf16
  call $log
)
"#, "trap" }
}
