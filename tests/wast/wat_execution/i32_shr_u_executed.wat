;; vybe-test: wast/wat_execution/i32_shr_u_executed
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
    i32.const 128
    i32.const 2
    i32.shr_u
    i32.const 32 call $vybe_check_i32))
