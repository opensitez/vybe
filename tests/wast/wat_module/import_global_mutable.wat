;; vybe-test: wast/wat_module/import_global_mutable
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "g" (global (mut i32))))
