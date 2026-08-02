;; vybe-test: wast/wat_simd_full_vector/i8x16_bitmask_gathers_high_bits
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result i32)
  v128.const i8x16 -1 0 -128 0 0 0 0 0 0 0 0 0 0 0 0 -1
  i8x16.bitmask))
(assert_return (invoke "f") (i32.const 0))
