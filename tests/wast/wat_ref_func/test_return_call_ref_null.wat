;; vybe-test: wast/wat_ref_func/test_return_call_ref_null
;; origin: languages/wast/tests/wast/test_wat_ref_func.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $sig (func (result i32)))
(func (export "_start") (local $r (ref null $sig))
  ref.null $sig
  local.set $r
  
  local.get $r
  return_call_ref $sig
)
)
