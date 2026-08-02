;; vybe-test: wast/wat_simd_full_vector/extend_high_signed
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 0 0 0 0 0 0 0 0 -1 -2 3 4 5 6 7 8
  i16x8.extend_high_i8x16_s))
(assert_return (invoke "f") (v128.const i16x8 -1 -2 3 4 5 6 7 8))
