;; vybe-test: wast/wat_execution/folded_add_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (call $log (i32.add (i32.const 19) (i32.const 23)))))
