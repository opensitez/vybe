;; vybe-test: wast/wat_execution/folded_nested_if_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (call $log
      (if (result i32) (i32.const 1)
        (then (i32.const 100))
        (else (i32.const 200))))))
