;; vybe-test: wast/wat_execution_extended/test_float_rounding_ceil
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const -1.5
    f64.ceil
    call $log))
