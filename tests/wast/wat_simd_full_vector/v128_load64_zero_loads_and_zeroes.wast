;; vybe-test: wast/wat_simd_full_vector/v128_load64_zero_loads_and_zeroes
;; origin: languages/wast/tests/wast/test_wat_simd_full_vector.rs
;; vybe-test-mode: compile

(module
  (memory 1) (data (i32.const 0) "\09\00\00\00\00\00\00\00")
  (func (export "f") (result v128)
    i32.const 0 v128.load64_zero))
(assert_return (invoke "f") (v128.const i64x2 9 0))
