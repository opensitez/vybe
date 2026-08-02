;; vybe-test: wast/wat_control_flow/loop_with_br_continue
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) (local i32) (loop $l local.get 0 local.get 1 i32.add local.set 1 local.get 0 i32.const 1 i32.sub local.tee 0 br_if $l)))
