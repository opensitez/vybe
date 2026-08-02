;; vybe-test: wast/wat_float_conversions/test_i64_reinterpret_f64
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  f64.const 1.0
  i64.reinterpret_f64
  call $log_i64
)
)
