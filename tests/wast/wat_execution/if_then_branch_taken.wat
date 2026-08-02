;; vybe-test: wast/wat_execution/if_then_branch_taken
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1
    if
      i32.const 111
      call $log
    end))
