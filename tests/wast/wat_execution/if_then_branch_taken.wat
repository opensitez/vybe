;; vybe-test: wast/wat_execution/if_then_branch_taken
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
    if
      i32.const 111
      i32.const 111 call $vybe_check_i32
    end))
