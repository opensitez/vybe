;; vybe-test: wast/wat_module/import_func
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "log" (func (param i32))))
