;; vybe-test: wast/wat_execution/named_param_used_in_body
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func $square (param $n i32) (result i32)
    local.get $n local.get $n i32.mul)
  (func (export "_start")
    i32.const 9
    call $square
    i32.const 81 call $vybe_check_i32))
