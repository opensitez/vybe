;; vybe-test: wast/wat_execution_extended/test_i32_rem_u_negative_operand
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const -5
    i32.const 3
    i32.rem_u
    call $log))
