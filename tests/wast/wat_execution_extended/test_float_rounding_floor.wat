;; vybe-test: wast/wat_execution_extended/test_float_rounding_floor
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

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
    f64.const -1.2
    f64.floor
    i32.const -2 call $vybe_check_i32))
