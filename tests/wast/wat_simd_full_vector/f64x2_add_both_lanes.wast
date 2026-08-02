;; vybe-test: wast/wat_simd_full_vector/f64x2_add_both_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f64)
  v128.const f64x2 1.5 2.5
  v128.const f64x2 0.5 0.5
  f64x2.add
  f64x2.extract_lane 0)
  (func (export "f1") (result f64)
  v128.const f64x2 1.5 2.5
  v128.const f64x2 0.5 0.5
  f64x2.add
  f64x2.extract_lane 1)
)
(assert_return (invoke "f0") (f64.const 2.0))
(assert_return (invoke "f1") (f64.const 3.0))
