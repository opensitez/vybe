;; vybe-test: wast/wat_simd_full_vector/shuffle_interleave_low
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15
  v128.const i8x16 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31
  i8x16.shuffle 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23))
(assert_return (invoke "f") (v128.const i8x16 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23))
