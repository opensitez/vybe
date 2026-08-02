;; vybe-test: wast/wat_execution_extended/test_conversion_trunc_trap
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs
;; vybe-test-mode: compile

(module
  (func (export "run") (result i32)
    f32.const 3e10
    i32.trunc_f32_s))
(assert_trap (invoke "run") "invalid conversion to integer")
