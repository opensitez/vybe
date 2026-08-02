;; vybe-test: wast/wat_simd_full_vector/i32x4_eq_mask
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 1 2 3 3
  v128.const i32x4 1 9 3 3
  i32x4.eq))
(assert_return (invoke "f") (v128.const i32x4 -1 0 -1 -1))
