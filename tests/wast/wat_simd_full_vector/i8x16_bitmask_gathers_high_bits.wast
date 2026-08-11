;; vybe-test: wast/wat_simd_full_vector/i8x16_bitmask_gathers_high_bits
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run
;; Spec (proposals/*/proposals/simd/SIMD.md, "Bitmask extraction"):
;; "These operations extract the high bit for each lane in `a` and produce a
;; scalar mask with all bits concatenated." Lanes 0 (-1), 2 (-128), and
;; 15 (-1) have their high bit set -> 1 | 4 | 32768 = 32773. The previous
;; expectation of 0 was an extraction-captured wrong value, not the spec.

(module (func (export "f") (result i32)
  v128.const i8x16 -1 0 -128 0 0 0 0 0 0 0 0 0 0 0 0 -1
  i8x16.bitmask))
(assert_return (invoke "f") (i32.const 32773))
