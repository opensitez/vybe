;; vybe-test: wast/wast_script/assert_trap_unreachable
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "f") unreachable))
(assert_trap (invoke "f") "unreachable")
