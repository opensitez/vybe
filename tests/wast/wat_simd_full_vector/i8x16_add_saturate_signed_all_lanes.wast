;; vybe-test: wast/wat_simd_full_vector/i8x16_add_saturate_signed_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 127 127 -128 -128 1 2 3 4 5 6 7 8 9 10 11 12
  v128.const i8x16 1 127 -1 -128 1 1 1 1 1 1 1 1 1 1 1 1
  i8x16.add_sat_s))
(assert_return (invoke "f") (v128.const i8x16 127 127 -128 -128 2 3 4 5 6 7 8 9 10 11 12 13))
