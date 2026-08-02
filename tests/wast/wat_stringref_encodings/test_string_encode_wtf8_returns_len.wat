;; vybe-test: wast/wat_stringref_encodings/test_string_encode_wtf8_returns_len
;; origin: languages/wast/tests/wast/test_wat_stringref_encodings.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\41\42\43")
(func (export "_start")
  i32.const 0 i32.const 3 string.new_utf8
  i32.const 10 string.encode_wtf8
  call $log)
)
