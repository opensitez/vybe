;; vybe-test: wast/wat_call_indirect/test_call_indirect_signature_mismatch_params
;; origin: languages/wast/tests/wast/test_wat_call_indirect.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $sig1 (func (result i32)))
(type $sig2 (func (param i32) (result i32)))
(table 1 funcref)
(func $f1 (type $sig2) 
  local.get 0)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  call_indirect (type $sig1) ;; calling a func that takes 1 param as if it takes 0
  call $log
)
)
