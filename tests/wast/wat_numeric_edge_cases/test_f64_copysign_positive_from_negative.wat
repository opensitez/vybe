;; vybe-test: wast/wat_numeric_edge_cases/test_f64_copysign_positive_from_negative
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        f64.const -3.0 f64.const 1.0 f64.copysign call $log_f64)
)
