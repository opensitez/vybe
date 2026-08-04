;; vybe-test: wast/wat_execution/multiple_log_calls_in_order
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
    i32.const 1 i32.const 1 call $vybe_check_i32
    i32.const 2 i32.const 2 call $vybe_check_i32
    i32.const 3 i32.const 3 call $vybe_check_i32
    i32.const 4 i32.const 4 call $vybe_check_i32
    i32.const 5 i32.const 5 call $vybe_check_i32))
