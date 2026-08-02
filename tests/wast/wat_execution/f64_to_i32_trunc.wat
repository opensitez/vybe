;; vybe-test: wast/wat_execution/f64_to_i32_trunc
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    f64.const 3.9
    i32.trunc_f64_s
    call $log))
