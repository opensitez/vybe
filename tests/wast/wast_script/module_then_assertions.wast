;; vybe-test: wast/wast_script/module_then_assertions
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module
  (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add)
  (func (export "mul") (param i32 i32) (result i32) local.get 0 local.get 1 i32.mul)
)
(assert_return (invoke "add" (i32.const 2) (i32.const 3)) (i32.const 5))
(assert_return (invoke "mul" (i32.const 4) (i32.const 5)) (i32.const 20))
(assert_return (invoke "add" (i32.const 0) (i32.const 0)) (i32.const 0))
