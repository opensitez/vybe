;; vybe-test: wast/wat_number_literals/test_float_exponent_positive
;; origin: languages/wast/tests/wast/test_wat_number_literals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") f64.const 1.5e3 call $log_f64)
)
