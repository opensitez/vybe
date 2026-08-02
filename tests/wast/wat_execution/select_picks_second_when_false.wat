;; vybe-test: wast/wat_execution/select_picks_second_when_false
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 10
    i32.const 20
    i32.const 0
    select
    call $log))
