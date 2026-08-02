;; vybe-test: wast/wat_simd_full_vector/f32x4_nearest_ties_to_even
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f32)
  v128.const f32x4 0.5 1.5 2.5 3.5
  f32x4.nearest
  f32x4.extract_lane 0)
  (func (export "f1") (result f32)
  v128.const f32x4 0.5 1.5 2.5 3.5
  f32x4.nearest
  f32x4.extract_lane 1)
  (func (export "f2") (result f32)
  v128.const f32x4 0.5 1.5 2.5 3.5
  f32x4.nearest
  f32x4.extract_lane 2)
  (func (export "f3") (result f32)
  v128.const f32x4 0.5 1.5 2.5 3.5
  f32x4.nearest
  f32x4.extract_lane 3)
)
(assert_return (invoke "f0") (f32.const 0.0))
(assert_return (invoke "f1") (f32.const 2.0))
(assert_return (invoke "f2") (f32.const 2.0))
(assert_return (invoke "f3") (f32.const 4.0))
