;; vybe-test: wast/wat_simd_full_vector/i32x4_shr_s_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 16 -16 7 -7
  i32.const 1
  i32x4.shr_s))
(assert_return (invoke "f") (v128.const i32x4 8 -8 3 -4))
