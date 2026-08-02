;; vybe-test: wast/wat_string_encode/test_string_encode_utf8_oob
;; origin: languages/wast/tests/wast/test_wat_string_encode.rs
;; vybe-test-mode: run-fail

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
  
  i32.const 65535
  string.encode_utf8
  drop
  
  i32.const 42
  call $log
)
)
