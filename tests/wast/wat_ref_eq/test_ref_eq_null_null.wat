;; vybe-test: wast/wat_ref_eq/test_ref_eq_null_null
;; origin: languages/wast/tests/wast/test_wat_ref_eq.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null func
  ref.null func
  ref.eq
  call $log
)
)
