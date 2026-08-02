;; vybe-test: wast/wat_simd_full_vector/i32x4_ge_s_mask
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 5 5 5 5
  v128.const i32x4 5 6 4 -1
  i32x4.ge_s))
(assert_return (invoke "f") (v128.const i32x4 -1 0 -1 -1))
