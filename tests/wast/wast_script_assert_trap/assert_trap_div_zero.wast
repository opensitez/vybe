;; vybe-test: wast/wast_script_assert_trap/assert_trap_div_zero
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) i32.const 1 i32.const 0 i32.div_s))
(assert_trap (invoke "f") "integer divide by zero")
