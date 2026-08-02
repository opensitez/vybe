;; vybe-test: wast/wat_simd_full_vector/narrow_i32x4_unsigned_saturates
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i32x4 -1 0 65535 65536
  v128.const i32x4 100000 -100000 1 2
  i16x8.narrow_i32x4_u))
(assert_return (invoke "f") (v128.const i16x8 0 0 65535 65535 65535 0 1 2))
