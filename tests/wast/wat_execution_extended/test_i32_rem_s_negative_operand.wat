;; vybe-test: wast/wat_execution_extended/test_i32_rem_s_negative_operand
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

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
    i32.const -5
    i32.const 3
    i32.rem_s
    i32.const -2 call $vybe_check_i32))
