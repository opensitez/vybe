;; vybe-test: wast/wat_control_flow/br_to_loop
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (loop $l br $l)))
