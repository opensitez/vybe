;; vybe-test: wast/wat_float_conversions/test_i32_trunc_f32_u_trap_neg
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  f32.const -1.0
  i32.trunc_f32_u
  call $log
)
)
