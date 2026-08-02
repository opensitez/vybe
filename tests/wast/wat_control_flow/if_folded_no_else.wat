;; vybe-test: wast/wat_control_flow/if_folded_no_else
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) (if (local.get 0) (then nop))))
