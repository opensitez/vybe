;; vybe-test: wast/wat_simd_full_vector/f64x2_sqrt_with_nan_lane
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module
  (func (export "f0") (result f64)
  v128.const f64x2 16.0 -4.0
  f64x2.sqrt
  f64x2.extract_lane 0)
  (func (export "f1") (result f64)
  v128.const f64x2 16.0 -4.0
  f64x2.sqrt
  f64x2.extract_lane 1)
)
(assert_return (invoke "f0") (f64.const 4.0))
(assert_return (invoke "f1") (f64.const nan:canonical))
