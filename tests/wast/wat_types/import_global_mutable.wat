;; vybe-test: wast/wat_types/import_global_mutable
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (import "env" "g1" (global (mut i32))))
