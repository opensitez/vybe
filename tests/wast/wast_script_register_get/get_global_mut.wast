;; vybe-test: wast/wast_script_register_get/get_global_mut
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: run

(module 
  (global (export "g") (mut i32) (i32.const 42))
  (func (export "set") (param i32) local.get 0 global.set 0)
)
(assert_return (get "g") (i32.const 42))
(invoke "set" (i32.const 99))
(assert_return (get "g") (i32.const 99))
