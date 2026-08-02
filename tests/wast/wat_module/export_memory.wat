;; vybe-test: wast/wat_module/export_memory
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (memory 1) (export "mem" (memory 0)))
