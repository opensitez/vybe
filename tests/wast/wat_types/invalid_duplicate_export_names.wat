;; vybe-test: wast/wat_types/invalid_duplicate_export_names
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile-fail

(module (func $f1) (func $f2) (export "f" (func $f1)) (export "f" (func $f2)))
