;; vybe-test: wast/wast_script/get_global_from_named_module
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module $m (global (export "g") i32 (i32.const 7)))
(assert_return (get $m "g") (i32.const 7))
