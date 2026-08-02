;; vybe-test: wast/wat_simd_full_vector/i8x16_all_true_true_when_all_nonzero
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result i32)
  v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16
  i8x16.all_true))
(assert_return (invoke "f") (i32.const 1))
