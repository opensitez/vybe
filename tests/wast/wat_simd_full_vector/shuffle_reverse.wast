;; vybe-test: wast/wat_simd_full_vector/shuffle_reverse
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15
  v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  i8x16.shuffle 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0))
(assert_return (invoke "f") (v128.const i8x16 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0))
