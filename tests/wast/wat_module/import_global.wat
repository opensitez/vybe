;; vybe-test: wast/wat_module/import_global
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "g" (global i32)))
