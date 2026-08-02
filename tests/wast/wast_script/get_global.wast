;; vybe-test: wast/wast_script/get_global
;; origin: languages/wast/tests/wast/test_wast_script.rs
;; vybe-test-mode: compile

(module (global (export "g") i32 (i32.const 42)))
(assert_return (get "g") (i32.const 42))
