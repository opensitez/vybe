;; vybe-test: wast/wat_f64_rounding/test_f64_nearest_neg_half_even
;; origin: languages/wast/tests/wast/test_wat_f64_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f64.const -1.5 f64.nearest call $log_f64)
)
