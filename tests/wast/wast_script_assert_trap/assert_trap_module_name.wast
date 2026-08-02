;; vybe-test: wast/wast_script_assert_trap/assert_trap_module_name
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module $m (func (export "f") unreachable))
(assert_trap (invoke $m "f") "unreachable")
