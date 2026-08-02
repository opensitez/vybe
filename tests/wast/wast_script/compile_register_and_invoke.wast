;; vybe-test: wast/wast_script/compile_register_and_invoke
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module $lib (func (export "double") (param i32) (result i32) local.get 0 i32.const 2 i32.mul))
(register "lib")
(invoke "double" (i32.const 5))
