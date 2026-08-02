;; vybe-test: wast/wat_string_eq/test_string_eq_diff_length
;; origin: languages/wast/tests/wast/test_wat_string_eq.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
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
)
