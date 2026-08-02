;; vybe-test: wast/wat_execution/named_param_used_in_body
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $square (param $n i32) (result i32)
    local.get $n local.get $n i32.mul)
  (func (export "_start")
    i32.const 9
    call $square
    call $log))
