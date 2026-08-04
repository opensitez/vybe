;; vybe-test: wast/wat_numeric_edge_cases/test_f64_min_with_nan_is_nan
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs

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
  (func (export "_start")
        f64.const 1.0 f64.const nan f64.min call $vybe_check_nan_f64)
)
