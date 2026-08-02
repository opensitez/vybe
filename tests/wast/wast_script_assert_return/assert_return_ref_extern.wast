;; vybe-test: wast/wast_script_assert_return/assert_return_ref_extern
;; origin: languages/wast/tests/wast/test_wast_script_assert_return.rs
;; vybe-test-mode: compile

(module
  (func (export "f") (param externref) (result externref) local.get 0)
)
(assert_return (invoke "f" (ref.extern 1)) (ref.extern 1))
