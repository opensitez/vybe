;; vybe-test: wast/wat_simd_full_vector/f32x4_max_nan_propagates_per_lane
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f32)
  v128.const f32x4 nan 5.0 3.0 7.0
  v128.const f32x4 1.0 nan 8.0 2.0
  f32x4.max
  f32x4.extract_lane 0)
  (func (export "f1") (result f32)
  v128.const f32x4 nan 5.0 3.0 7.0
  v128.const f32x4 1.0 nan 8.0 2.0
  f32x4.max
  f32x4.extract_lane 1)
  (func (export "f2") (result f32)
  v128.const f32x4 nan 5.0 3.0 7.0
  v128.const f32x4 1.0 nan 8.0 2.0
  f32x4.max
  f32x4.extract_lane 2)
  (func (export "f3") (result f32)
  v128.const f32x4 nan 5.0 3.0 7.0
  v128.const f32x4 1.0 nan 8.0 2.0
  f32x4.max
  f32x4.extract_lane 3)
)
(assert_return (invoke "f0") (f32.const nan:canonical))
(assert_return (invoke "f1") (f32.const nan:canonical))
(assert_return (invoke "f2") (f32.const 8.0))
(assert_return (invoke "f3") (f32.const 7.0))
