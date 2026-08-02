;; vybe-test: wast/wat_control_flow/br_table_default
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (param i32) (block $a (block $b (block $c local.get 0 br_table $a $b $c)))))
