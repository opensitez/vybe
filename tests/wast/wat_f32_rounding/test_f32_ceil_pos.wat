;; vybe-test: wast/wat_f32_rounding/test_f32_ceil_pos
;; origin: languages/wast/tests/wast/test_wat_f32_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f32.const 1.2 f32.ceil call $log_f32)
)
