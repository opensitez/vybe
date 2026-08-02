;; vybe-test: wast/wat_stringref_encodings/test_string_new_wtf16_roundtrips
;; origin: languages/wast/tests/wast/test_wat_stringref_encodings.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\48\00\69\00")
(data (i32.const 10) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_wtf16
  i32.const 10 i32.const 2 string.new_utf8
  string.eq
  call $log)
)
