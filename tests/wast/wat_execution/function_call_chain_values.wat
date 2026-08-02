;; vybe-test: wast/wat_execution/function_call_chain_values
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $inc (param $x i32) (result i32) local.get $x i32.const 1 i32.add)
  (func $double (param $x i32) (result i32) local.get $x i32.const 2 i32.mul)
  (func (export "_start")
    i32.const 5
    call $inc
    call $log      ;; 6
    i32.const 5
    call $double
    call $log      ;; 10
    i32.const 5
    call $inc
    call $double
    call $log))    ;; 12
