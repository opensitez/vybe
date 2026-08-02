;; vybe-test: wast/wast_script/assert_trap_stack_overflow
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func $rec (export "rec") call $rec))
(assert_trap (invoke "rec") "call stack exhausted")
