;; vybe-test: wast/wat_simd_full_vector/i8x16_min_s_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 -1 5 -128 127 9 9 9 9 9 9 9 9 9 9 9 9
  v128.const i8x16 1 -5 127 -128 0 0 0 0 0 0 0 0 0 0 0 0
  i8x16.min_s))
(assert_return (invoke "f") (v128.const i8x16 -1 -5 -128 -128 0 0 0 0 0 0 0 0 0 0 0 0))
