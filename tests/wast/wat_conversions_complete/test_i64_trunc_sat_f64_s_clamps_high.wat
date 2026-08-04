;; vybe-test: wast/wat_conversions_complete/test_i64_trunc_sat_f64_s_clamps_high
;; origin: languages/wast/tests/wast/test_wat_conversions_complete.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
  (func (export "_start")
        f64.const 1e19 i64.trunc_sat_f64_s i64.const 9223372036854775807 call $vybe_check_i64)
)
