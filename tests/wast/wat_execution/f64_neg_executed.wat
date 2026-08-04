;; vybe-test: wast/wat_execution/f64_neg_executed
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
    f64.const 5.0
    f64.neg
    i32.const -5 call $vybe_check_i32))
