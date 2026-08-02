;; vybe-test: wast/wast_script_register_get/register_named_module
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module $lib (func (export "f") (result i32) i32.const 42))
(register "lib2" $lib)
(module 
  (import "lib2" "f" (func $f (result i32)))
  (func (export "g") (result i32) call $f)
)
(assert_return (invoke "g") (i32.const 42))
