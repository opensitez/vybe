;; vybe-test: wast/wast_script_assert_trap/assert_trap_invalid_conversion_to_integer_overflow
;; origin: languages/wast/tests/wast/test_wast_script_assert_trap.rs
;; vybe-test-mode: compile

(module (func (export "f") (result i32) f32.const 3000000000.0 i32.trunc_f32_s))
(assert_trap (invoke "f") "integer overflow")
