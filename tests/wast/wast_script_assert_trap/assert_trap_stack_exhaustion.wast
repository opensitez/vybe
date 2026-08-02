;; vybe-test: wast/wast_script_assert_trap/assert_trap_stack_exhaustion
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (func $rec (export "f") call $rec))
(assert_trap (invoke "f") "call stack exhausted")
