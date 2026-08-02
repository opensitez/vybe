;; vybe-test: wast/wat_simd_full_vector/i32x4_trunc_sat_f32x4_u_saturates
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: run

(module (func (export "f") (result v128)
  v128.const f32x4 3.9 -1.0 5000000000.0 100.5
  i32x4.trunc_sat_f32x4_u))
(assert_return (invoke "f") (v128.const i32x4 3 0 -1 100))
