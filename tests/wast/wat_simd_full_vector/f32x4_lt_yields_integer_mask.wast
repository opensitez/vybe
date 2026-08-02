;; vybe-test: wast/wat_simd_full_vector/f32x4_lt_yields_integer_mask
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const f32x4 -1.5 5.0 2.0 2.0
  v128.const f32x4 0.0 0.0 2.0 3.0
  f32x4.lt))
(assert_return (invoke "f") (v128.const i32x4 -1 0 0 -1))
