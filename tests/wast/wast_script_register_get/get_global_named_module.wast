;; vybe-test: wast/wast_script_register_get/get_global_named_module
;; origin: languages/wast/tests/wast/test_wast_script_register_get.rs
;; vybe-test-mode: run

(module $m (global (export "g") i32 (i32.const 42)))
(assert_return (get $m "g") (i32.const 42))
