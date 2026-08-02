;; vybe-test: wast/wat_simd_full_vector/narrow_i16x8_signed_saturates
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i16x8 0 127 128 -129 -1 1 200 -200
  v128.const i16x8 100 -100 127 -128 0 300 -300 5
  i8x16.narrow_i16x8_s))
(assert_return (invoke "f") (v128.const i8x16 0 127 127 -128 -1 1 127 -128 100 -100 127 -128 0 127 -128 5))
