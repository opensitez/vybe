;; vybe-test: wast/wat_control_flow/block_br_break
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module (func (block $b i32.const 1 drop br $b)))
