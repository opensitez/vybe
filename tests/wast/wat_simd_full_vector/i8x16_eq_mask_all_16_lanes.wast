;; vybe-test: wast/wat_simd_full_vector/i8x16_eq_mask_all_16_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 1 2 3 4 5 6 7 8 1 2 3 4 5 6 7 8
  v128.const i8x16 1 0 3 0 5 0 7 0 1 0 3 0 5 0 7 0
  i8x16.eq))
(assert_return (invoke "f") (v128.const i8x16 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0 -1 0))
