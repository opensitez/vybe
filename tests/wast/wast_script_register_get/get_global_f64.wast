;; vybe-test: wast/wast_script_register_get/get_global_f64
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: compile

(module (global (export "g") f64 (f64.const 2.718)))
(assert_return (get "g") (f64.const 2.718))
