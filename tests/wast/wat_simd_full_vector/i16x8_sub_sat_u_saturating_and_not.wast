;; vybe-test: wast/wat_simd_full_vector/i16x8_sub_sat_u_saturating_and_not
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const i16x8 1000 5 300 0 100 200 65535 8
  v128.const i16x8 1 10 300 5 50 100 1 8
  i16x8.sub_sat_u))
(assert_return (invoke "f") (v128.const i16x8 999 0 0 0 50 100 65534 0))
