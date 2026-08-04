;; vybe-test: wast/wat_execution/i32_bitwise_and_executed
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
    i32.const 0xFF
    i32.const 0x0F
    i32.and
    i32.const 15 call $vybe_check_i32))
