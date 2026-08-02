;; vybe-test: wast/wat_i64_arithmetic/test_i64_add_overflow
;; origin: languages/wast/tests/wast/test_wat_i64_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") i64.const 9223372036854775807 i64.const 1 i64.add call $log_i64)
)
