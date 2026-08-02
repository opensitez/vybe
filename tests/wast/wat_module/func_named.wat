;; vybe-test: wast/wat_module/func_named
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func $add (param $a i32) (param $b i32) (result i32) local.get $a local.get $b i32.add))
