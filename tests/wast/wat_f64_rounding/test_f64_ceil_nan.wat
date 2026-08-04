;; vybe-test: wast/wat_f64_rounding/test_f64_ceil_nan
;; origin: languages/wast/tests/wast/test_wat_f64_rounding.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_nan_f64 (param f64)
    local.get 0
    local.get 0
    f64.eq
    if
      unreachable
    end)
  (func (export "_start") f64.const nan f64.ceil call $vybe_check_nan_f64)
)
