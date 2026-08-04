;; vybe-test: wast/wat_execution_extended/test_conversion_trunc_sat_negative
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    f32.const -3e10
    i32.trunc_sat_f32_s
    i32.const -2147483648 call $vybe_check_i32))
