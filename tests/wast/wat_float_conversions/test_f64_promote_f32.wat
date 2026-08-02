;; vybe-test: wast/wat_float_conversions/test_f64_promote_f32
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  f32.const 42.5
  f64.promote_f32
  call $log_f64
)
)
