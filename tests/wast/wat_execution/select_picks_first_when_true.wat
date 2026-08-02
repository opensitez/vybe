;; vybe-test: wast/wat_execution/select_picks_first_when_true
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 10
    i32.const 20
    i32.const 1
    select
    call $log))
