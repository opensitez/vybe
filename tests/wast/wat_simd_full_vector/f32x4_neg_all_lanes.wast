;; vybe-test: wast/wat_simd_full_vector/f32x4_neg_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f32)
  v128.const f32x4 1.5 -2.5 3.5 -4.5
  f32x4.neg
  f32x4.extract_lane 0)
  (func (export "f1") (result f32)
  v128.const f32x4 1.5 -2.5 3.5 -4.5
  f32x4.neg
  f32x4.extract_lane 1)
  (func (export "f2") (result f32)
  v128.const f32x4 1.5 -2.5 3.5 -4.5
  f32x4.neg
  f32x4.extract_lane 2)
  (func (export "f3") (result f32)
  v128.const f32x4 1.5 -2.5 3.5 -4.5
  f32x4.neg
  f32x4.extract_lane 3)
)
(assert_return (invoke "f0") (f32.const -1.5))
(assert_return (invoke "f1") (f32.const 2.5))
(assert_return (invoke "f2") (f32.const -3.5))
(assert_return (invoke "f3") (f32.const 4.5))
