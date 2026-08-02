;; vybe-test: wast/wat_f32_rounding/test_f32_nearest_pos_half_even_2
;; origin: languages/wast/tests/wast/test_wat_f32_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f32.const 2.5 f32.nearest call $log_f32)
)
