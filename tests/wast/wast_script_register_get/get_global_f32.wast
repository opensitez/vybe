;; vybe-test: wast/wast_script_register_get/get_global_f32
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module (global (export "g") f32 (f32.const 3.14)))
(assert_return (get "g") (f32.const 3.14))
