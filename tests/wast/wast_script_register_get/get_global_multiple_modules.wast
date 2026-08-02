;; vybe-test: wast/wast_script_register_get/get_global_multiple_modules
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module $m1 (global (export "g") i32 (i32.const 42)))
(module $m2 (global (export "g") i32 (i32.const 99)))
(assert_return (get $m1 "g") (i32.const 42))
(assert_return (get $m2 "g") (i32.const 99))
