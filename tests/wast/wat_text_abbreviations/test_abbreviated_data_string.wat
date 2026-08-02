;; vybe-test: wast/wat_text_abbreviations/test_abbreviated_data_string
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a\00\00\00")
        (func (export "_start") (call $log (i32.load (i32.const 0)))))
