;; vybe-test: wast/wat_execution_extended/test_float_copysign_zero
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 1.0
    f64.const -0.0
    f64.copysign
    call $log))
