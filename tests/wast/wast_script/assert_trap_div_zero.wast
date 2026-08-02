;; vybe-test: wast/wast_script/assert_trap_div_zero
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (func (export "div") (param i32 i32) (result i32) local.get 0 local.get 1 i32.div_s))
(assert_trap (invoke "div" (i32.const 1) (i32.const 0)) "integer divide by zero")
