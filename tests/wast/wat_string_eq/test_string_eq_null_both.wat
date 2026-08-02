;; vybe-test: wast/wat_string_eq/test_string_eq_null_both
;; origin: languages/wast/tests/wast/test_wat_string_eq.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null string
  ref.null string
  string.eq
  call $log
)
)
