;; vybe-test: wast/wat_control_flow/if_with_label
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) local.get 0 if $l nop end))
