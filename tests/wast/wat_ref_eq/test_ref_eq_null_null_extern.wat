;; vybe-test: wast/wat_ref_eq/test_ref_eq_null_null_extern
;; origin: languages/wast/tests/wast/test_wat_ref_eq.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null extern
  ref.null extern
  ref.eq
  call $log
)
)
