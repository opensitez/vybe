;; vybe-test: wast/wat_execution/nop_is_transparent
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
    nop
    nop
    i32.const 5
    nop
    i32.const 5 call $vybe_check_i32))
