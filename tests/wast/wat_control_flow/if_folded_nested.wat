;; vybe-test: wast/wat_control_flow/if_folded_nested
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (func (param i32) (result i32)
    (if (result i32) (local.get 0)
      (then
        (if (result i32) (i32.const 1)
          (then (i32.const 10))
          (else (i32.const 20))))
      (else (i32.const 0)))))
