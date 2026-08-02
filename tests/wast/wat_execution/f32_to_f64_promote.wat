;; vybe-test: wast/wat_execution/f32_to_f64_promote
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f32.const 2.0
    f64.promote_f32
    call $log))
