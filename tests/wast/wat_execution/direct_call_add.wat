;; vybe-test: wast/wat_execution/direct_call_add
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $add (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.add)
  (func (export "_start")
    i32.const 13
    i32.const 29
    call $add
    call $log))
