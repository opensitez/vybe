;; vybe-test: wast/wat_f32_rounding/test_f32_floor_inf
;; origin: languages/wast/tests/wast/test_wat_f32_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f32.const inf f32.floor call $log_f32)
)
