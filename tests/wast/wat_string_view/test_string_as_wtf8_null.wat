;; vybe-test: wast/wat_string_view/test_string_as_wtf8_null
;; origin: languages/wast/tests/wast/test_wat_string_view.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null string
  string.as_wtf8
  drop
  i32.const 42
  call $log
)
)
