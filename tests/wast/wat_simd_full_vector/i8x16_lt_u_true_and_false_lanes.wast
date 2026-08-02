;; vybe-test: wast/wat_simd_full_vector/i8x16_lt_u_true_and_false_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 0 255 16 128 1 1 1 1 1 1 1 1 1 1 1 1
  v128.const i8x16 1 0 32 127 1 1 1 1 1 1 1 1 1 1 1 1
  i8x16.lt_u))
(assert_return (invoke "f") (v128.const i8x16 -1 0 -1 0 0 0 0 0 0 0 0 0 0 0 0 0))
