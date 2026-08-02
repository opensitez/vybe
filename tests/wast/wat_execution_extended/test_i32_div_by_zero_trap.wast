;; vybe-test: wast/wat_execution_extended/test_i32_div_by_zero_trap
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (func (export "run") (result i32)
    i32.const 42
    i32.const 0
    i32.div_s))
(assert_trap (invoke "run") "integer divide by zero")
