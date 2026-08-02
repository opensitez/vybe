;; vybe-test: wast/wat_simd_full_vector/i32x4_shl_all_lanes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 1 2 3 -1
  i32.const 4
  i32x4.shl))
(assert_return (invoke "f") (v128.const i32x4 16 32 48 -16))
