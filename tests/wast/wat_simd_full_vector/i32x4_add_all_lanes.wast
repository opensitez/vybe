;; vybe-test: wast/wat_simd_full_vector/i32x4_add_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 1 2 3 4
  v128.const i32x4 10 20 30 40
  i32x4.add))
(assert_return (invoke "f") (v128.const i32x4 11 22 33 44))
