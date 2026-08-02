;; vybe-test: wast/wat_simd_full_vector/i32x4_lt_u_true_and_false_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 1 -1 10 50
  v128.const i32x4 2 -2 10 999
  i32x4.lt_u))
(assert_return (invoke "f") (v128.const i32x4 -1 0 0 -1))
