;; vybe-test: wast/wast_script_register_get/get_global_i64
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module (global (export "g") i64 (i64.const 9999999999)))
(assert_return (get "g") (i64.const 9999999999))
