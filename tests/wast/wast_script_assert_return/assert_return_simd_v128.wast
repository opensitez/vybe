;; vybe-test: wast/wast_script_assert_return/assert_return_simd_v128
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module 
  (func (export "f") (result v128) v128.const i32x4 1 2 3 4)
)
(assert_return (invoke "f") (v128.const i32x4 1 2 3 4))
