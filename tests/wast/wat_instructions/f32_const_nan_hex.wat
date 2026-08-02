;; vybe-test: wast/wat_instructions/f32_const_nan_hex
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func (result f32) f32.const nan:0x200000))
