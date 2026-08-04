;; vybe-test: wast/wat_execution/i32_sub_negative_result
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
    i32.const 3
    i32.const 10
    i32.sub
    i32.const -7 call $vybe_check_i32))
