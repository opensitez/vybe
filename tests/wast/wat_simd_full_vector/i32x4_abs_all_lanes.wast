;; vybe-test: wast/wat_simd_full_vector/i32x4_abs_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 -5 5 -2147483648 0
  i32x4.abs))
(assert_return (invoke "f") (v128.const i32x4 5 5 -2147483648 0))
