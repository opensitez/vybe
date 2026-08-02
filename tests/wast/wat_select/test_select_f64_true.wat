;; vybe-test: wast/wat_select/test_select_f64_true
;; origin: languages/wast/tests/wast/test_wat_select.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  f64.const 3.14
  f64.const 2.71
  i32.const 42
  select
  call $log_f64
)
)
