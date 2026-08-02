;; vybe-test: wast/wat_conversions_complete/test_i32_trunc_sat_f32_u_negative_is_zero
;; origin: languages/wast/tests/wast/test_wat_conversions_complete.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        f32.const -3.0 i32.trunc_sat_f32_u call $log)
)
