;; vybe-test: wast/wat_text_abbreviations/test_implicit_type_from_params
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $add (param i32 i32) (result i32) (i32.add (local.get 0) (local.get 1)))
        (func (export "_start") (call $log (call $add (i32.const 40) (i32.const 2)))))
