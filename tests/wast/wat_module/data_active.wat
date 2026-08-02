;; vybe-test: wast/wat_module/data_active
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (memory 1) (data (offset (i32.const 0)) "hello"))
