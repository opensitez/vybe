;; vybe-test: wast/wat_simd_full_vector/i8x16_sub_sat_u_saturating_and_not
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 10 5 200 100 50 0 255 8 1 2 3 4 5 6 7 8
  v128.const i8x16 3 10 50 100 0 5 1 8 1 2 3 4 5 6 7 8
  i8x16.sub_sat_u))
(assert_return (invoke "f") (v128.const i8x16 7 0 150 0 50 0 254 0 0 0 0 0 0 0 0 0))
