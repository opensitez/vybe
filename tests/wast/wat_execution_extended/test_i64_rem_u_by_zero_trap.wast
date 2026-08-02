;; vybe-test: wast/wat_execution_extended/test_i64_rem_u_by_zero_trap
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (func (export "run") (result i64)
    i64.const 42
    i64.const 0
    i64.rem_u))
(assert_trap (invoke "run") "integer divide by zero")
