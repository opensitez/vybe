;; vybe-test: wast/wat_simd_full_vector/avgr_u_rounds_up_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 255 1 10 100 0 2 4 6 8 10 12 14 16 18 20 22
  v128.const i8x16 255 2 13 101 0 2 4 6 8 10 12 14 16 18 20 22
  i8x16.avgr_u))
(assert_return (invoke "f") (v128.const i8x16 255 2 12 101 0 2 4 6 8 10 12 14 16 18 20 22))
