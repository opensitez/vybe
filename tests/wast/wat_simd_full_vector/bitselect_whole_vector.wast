;; vybe-test: wast/wat_simd_full_vector/bitselect_whole_vector
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 -1 -1 -1 -1
  v128.const i32x4 0 0 0 0
  v128.const i32x4 -1 0 -1 0
  v128.bitselect))
(assert_return (invoke "f") (v128.const i32x4 -1 0 -1 0))
