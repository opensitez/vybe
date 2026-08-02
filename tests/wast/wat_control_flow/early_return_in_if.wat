;; vybe-test: wast/wat_control_flow/early_return_in_if
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (func (export "f") (param i32) (result i32)
    local.get 0
    if (result i32)
      i32.const 1
      return
    end
    i32.const 0))
