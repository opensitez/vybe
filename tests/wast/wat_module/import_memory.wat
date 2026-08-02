;; vybe-test: wast/wat_module/import_memory
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "mem" (memory 1)))
