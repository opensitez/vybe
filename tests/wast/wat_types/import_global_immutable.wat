;; vybe-test: wast/wat_types/import_global_immutable
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (import "env" "g1" (global i32)))
