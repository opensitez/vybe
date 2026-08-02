;; vybe-test: wast/wat_stringref_views/test_wtf16_encode_count
;; origin: languages/wast/tests/wast/test_wat_stringref_views.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start")
  i32.const 0 i32.const 2 string.new_utf8
  string.as_wtf16
  i32.const 10 i32.const 0 i32.const 2
  stringview_wtf16.encode
  call $log)
)
