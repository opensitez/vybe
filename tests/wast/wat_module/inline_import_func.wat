;; vybe-test: wast/wat_module/inline_import_func
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func $log (import "env" "log") (param i32)))
