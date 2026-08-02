;; vybe-test: wast/wat_simd_full_vector/dot_product_pairs
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i16x8 1 2 3 4 5 6 7 8
  v128.const i16x8 1 1 2 2 3 3 4 4
  i32x4.dot_i16x8_s))
(assert_return (invoke "f") (v128.const i32x4 3 14 33 60))
