;; vybe-test: wast/wat_simd_full_vector/swizzle_gather_and_zero
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i8x16 100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115
  v128.const i8x16 15 0 3 16 1 1 200 7 8 9 10 11 12 13 14 2
  i8x16.swizzle))
(assert_return (invoke "f") (v128.const i8x16 115 100 103 0 101 101 0 107 108 109 110 111 112 113 114 102))
