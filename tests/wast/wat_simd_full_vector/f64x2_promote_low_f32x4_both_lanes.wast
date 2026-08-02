;; vybe-test: wast/wat_simd_full_vector/f64x2_promote_low_f32x4_both_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f64)
  v128.const f32x4 1.5 2.5 9.0 9.0
  f64x2.promote_low_f32x4
  f64x2.extract_lane 0)
  (func (export "f1") (result f64)
  v128.const f32x4 1.5 2.5 9.0 9.0
  f64x2.promote_low_f32x4
  f64x2.extract_lane 1)
)
(assert_return (invoke "f0") (f64.const 1.5))
(assert_return (invoke "f1") (f64.const 2.5))
