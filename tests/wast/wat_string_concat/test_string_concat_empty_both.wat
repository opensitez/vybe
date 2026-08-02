;; vybe-test: wast/wat_string_concat/test_string_concat_empty_both
;; origin: languages/wast/tests/wast/test_wat_string_concat.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
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
)
