;; vybe-test: wast/wast_script_assert_trap/assert_trap_uninitialized_element
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (type $t (func)) (table 1 funcref) (func (export "f") i32.const 0 call_indirect (type $t)))
(assert_trap (invoke "f") "uninitialized element")
