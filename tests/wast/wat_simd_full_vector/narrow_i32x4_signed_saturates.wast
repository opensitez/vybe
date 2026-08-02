;; vybe-test: wast/wat_simd_full_vector/narrow_i32x4_signed_saturates
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 0 32767 32768 -32769
  v128.const i32x4 -40000 40000 -1 1
  i16x8.narrow_i32x4_s))
(assert_return (invoke "f") (v128.const i16x8 0 32767 32767 -32768 -32768 32767 -1 1))
