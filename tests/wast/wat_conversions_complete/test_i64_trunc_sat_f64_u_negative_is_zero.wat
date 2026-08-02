;; vybe-test: wast/wat_conversions_complete/test_i64_trunc_sat_f64_u_negative_is_zero
;; origin: languages/wast/tests/wast/test_wat_conversions_complete.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        f64.const -1.0 i64.trunc_sat_f64_u call $log_i64)
)
