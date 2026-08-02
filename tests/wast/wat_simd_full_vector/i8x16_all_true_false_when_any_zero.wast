;; vybe-test: wast/wat_simd_full_vector/i8x16_all_true_false_when_any_zero
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result i32)
  v128.const i8x16 1 1 1 1 1 1 1 0 1 1 1 1 1 1 1 1
  i8x16.all_true))
(assert_return (invoke "f") (i32.const 0))
