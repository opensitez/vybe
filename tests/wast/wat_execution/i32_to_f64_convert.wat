;; vybe-test: wast/wat_execution/i32_to_f64_convert
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
    i32.const 7
    f64.convert_i32_s
    i32.const 7 call $vybe_check_i32))
