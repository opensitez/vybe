;; vybe-test: wast/wat_float_conversions/test_i32_trunc_sat_f32_s_pos
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  f32.const 3000000000.0 ;; > i32.max
  i32.trunc_sat_f32_s
  call $log
)
)
