;; vybe-test: wast/wat_instructions/local_get_by_name
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (param $x i32) (result i32) local.get $x))
