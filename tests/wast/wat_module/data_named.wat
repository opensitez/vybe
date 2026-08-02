;; vybe-test: wast/wat_module/data_named
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (memory 1) (data $d (offset (i32.const 0)) "world"))
