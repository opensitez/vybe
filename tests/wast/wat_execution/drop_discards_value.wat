;; vybe-test: wast/wat_execution/drop_discards_value
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 999
    drop
    i32.const 1
    call $log))
