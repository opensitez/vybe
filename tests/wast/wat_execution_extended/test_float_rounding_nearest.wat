;; vybe-test: wast/wat_execution_extended/test_float_rounding_nearest
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f32)))
  (func (export "_start")
    f32.const 1.5
    f32.nearest
    call $log))
