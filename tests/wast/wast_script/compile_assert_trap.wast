;; vybe-test: wast/wast_script/compile_assert_trap
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "boom") unreachable))
(assert_trap (invoke "boom") "unreachable")
