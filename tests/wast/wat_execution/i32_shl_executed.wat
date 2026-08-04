;; vybe-test: wast/wat_execution/i32_shl_executed
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
    i32.const 1
    i32.const 3
    i32.shl
    i32.const 8 call $vybe_check_i32))
