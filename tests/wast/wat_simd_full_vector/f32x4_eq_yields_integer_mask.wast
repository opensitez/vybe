;; vybe-test: wast/wat_simd_full_vector/f32x4_eq_yields_integer_mask
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const f32x4 1.0 2.0 3.0 3.0
  v128.const f32x4 1.0 9.0 3.0 3.0
  f32x4.eq))
(assert_return (invoke "f") (v128.const i32x4 -1 0 -1 -1))
