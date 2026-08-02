;; vybe-test: wast/wat_execution/i32_to_f64_convert
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    i32.const 7
    f64.convert_i32_s
    call $log))
