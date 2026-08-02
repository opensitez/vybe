;; vybe-test: wast/wat_execution/if_else_false_branch
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0
    if
      i32.const 1
      call $log
    else
      i32.const 2
      call $log
    end))
