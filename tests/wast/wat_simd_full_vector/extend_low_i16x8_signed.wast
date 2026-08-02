;; vybe-test: wast/wat_simd_full_vector/extend_low_i16x8_signed
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i16x8 -1 -32768 32767 5 9 9 9 9
  i32x4.extend_low_i16x8_s))
(assert_return (invoke "f") (v128.const i32x4 -1 -32768 32767 5))
