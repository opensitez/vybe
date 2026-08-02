;; vybe-test: wast/wat_module/import_func_named
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "log" (func $log (param i32))))
