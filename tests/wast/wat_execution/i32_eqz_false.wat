;; vybe-test: wast/wat_execution/i32_eqz_false
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 5
    i32.eqz
    call $log))
