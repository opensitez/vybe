;; vybe-test: wast/wat_execution_extended/test_i64_div_s_overflow_trap
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (func (export "run") (result i64)
    i64.const -9223372036854775808
    i64.const -1
    i64.div_s))
(assert_trap (invoke "run") "integer overflow")
