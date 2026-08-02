;; vybe-test: wast/wat_execution_extended/test_conversion_trunc_sat
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    f32.const 3e10
    i32.trunc_sat_f32_s
    call $log))
