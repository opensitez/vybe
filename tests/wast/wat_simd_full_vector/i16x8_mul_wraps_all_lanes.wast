;; vybe-test: wast/wat_simd_full_vector/i16x8_mul_wraps_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i16x8 256 256 -1 32767 2 3 4 5
  v128.const i16x8 256 128 -1 2 2 3 4 5
  i16x8.mul))
(assert_return (invoke "f") (v128.const i16x8 0 -32768 1 -2 4 9 16 25))
