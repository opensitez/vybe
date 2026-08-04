;; vybe-test: wast/wat_execution/i32_mul_executed
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
  (func (export "_start")
    i32.const 6
    i32.const 7
    i32.mul
    i32.const 42 call $vybe_check_i32))
