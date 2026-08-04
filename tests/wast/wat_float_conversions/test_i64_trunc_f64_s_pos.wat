;; vybe-test: wast/wat_float_conversions/test_i64_trunc_f64_s_pos
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

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
  f64.const 42.9
  i64.trunc_f64_s
  i64.const 42 call $vybe_check_i64
)
)
