;; vybe-test: wast/wat_i64_arithmetic/test_i64_mul_overflow
;; origin: languages/wast/tests/wast/test_wat_i64_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") i64.const 3037000499 i64.const 3037000499 i64.mul call $log_i64)
)
