;; vybe-test: wast/wat_ref_func/test_ref_func_call_indirect
;; origin: languages/wast/tests/wast/test_wat_ref_func.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $sig (func (result i32)))
(table 1 funcref)
(func $f1 (result i32) i32.const 42)
(func (export "_start")
  i32.const 0
  ref.func $f1
  table.set 0
  
  i32.const 0
  call_indirect (type $sig)
  call $log
)
)
