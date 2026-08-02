;; vybe-test: wast/wat_simd_full_vector/extadd_pairwise_unsigned
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 255 1 10 20 0 0 100 100 1 2 3 4 5 6 7 8
  i16x8.extadd_pairwise_i8x16_u))
(assert_return (invoke "f") (v128.const i16x8 256 30 0 200 3 7 11 15))
