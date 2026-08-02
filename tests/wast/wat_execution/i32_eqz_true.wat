;; vybe-test: wast/wat_execution/i32_eqz_true
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0
    i32.eqz
    call $log))
