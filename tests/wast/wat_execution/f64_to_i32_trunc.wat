;; vybe-test: wast/wat_execution/f64_to_i32_trunc
;; origin: languages/wast/tests/wast/test_wat_execution.rs

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
    f64.const 3.9
    i32.trunc_f64_s
    i32.const 3 call $vybe_check_i32))
