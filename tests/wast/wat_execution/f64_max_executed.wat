;; vybe-test: wast/wat_execution/f64_max_executed
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
    f64.const 3.0
    f64.const 8.0
    f64.max
    i32.const 8 call $vybe_check_i32))
