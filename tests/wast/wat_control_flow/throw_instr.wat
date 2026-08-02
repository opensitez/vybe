;; vybe-test: wast/wat_control_flow/throw_instr
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (tag $e (param i32))
  (func (export "f") i32.const 42 throw $e))
