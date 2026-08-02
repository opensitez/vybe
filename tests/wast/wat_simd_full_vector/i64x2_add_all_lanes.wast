;; vybe-test: wast/wat_simd_full_vector/i64x2_add_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i64x2 1000000000000 -5
  v128.const i64x2 1 5
  i64x2.add))
(assert_return (invoke "f") (v128.const i64x2 1000000000001 0))
