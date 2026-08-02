;; vybe-test: wast/wat_types/export_memory
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (memory $m 1) (export "mem" (memory $m)))
