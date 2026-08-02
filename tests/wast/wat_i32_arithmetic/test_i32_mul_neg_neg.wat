;; vybe-test: wast/wat_i32_arithmetic/test_i32_mul_neg_neg
;; origin: languages/wast/tests/wast/test_wat_i32_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") i32.const -10 i32.const -20 i32.mul call $log)
)
