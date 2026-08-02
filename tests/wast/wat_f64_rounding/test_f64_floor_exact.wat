;; vybe-test: wast/wat_f64_rounding/test_f64_floor_exact
;; origin: languages/wast/tests/wast/test_wat_f64_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f64.const 2.0 f64.floor call $log_f64)
)
