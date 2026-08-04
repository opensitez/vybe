;; vybe-test: wast/wat_execution/f64_sqrt_executed
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    f64.const 9.0
    f64.sqrt
    i32.const 3 call $vybe_check_i32))
