;; vybe-test: wast/wat_stringref_iter/test_iter_advance_clamps
;; origin: languages/wast/tests/wast/test_wat_stringref_iter.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\48\65\6C\6C\6F")
(func (export "_start")
  i32.const 0 i32.const 5 string.new_utf8
  string.as_iter
  i32.const 99
  stringview_iter.advance
  call $log)
)
