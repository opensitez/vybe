;; vybe-test: wast/wast_script_register_get/register_anonymous_module
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module (func (export "answer") (result i32) i32.const 42))
(register "lib")
(module
  (import "lib" "answer" (func $a (result i32)))
  (func (export "use") (result i32) call $a call $a i32.add))
(assert_return (invoke "use") (i32.const 84))
