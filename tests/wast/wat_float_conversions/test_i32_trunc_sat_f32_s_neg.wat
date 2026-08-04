;; vybe-test: wast/wat_float_conversions/test_i32_trunc_sat_f32_s_neg
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
  f32.const -3000000000.0 ;; < i32.min
  i32.trunc_sat_f32_s
  i32.const -2147483648 call $vybe_check_i32
)
)
