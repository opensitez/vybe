;; vybe-test: wast/wat_instructions/compile_add_func
;; origin: languages/wast/tests/wast/test_wat_instructions.rs
;; vybe-test-mode: compile

(module (func $add (export "add") (param $a i32) (param $b i32) (result i32) local.get $a local.get $b i32.add))
