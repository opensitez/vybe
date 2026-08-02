;; vybe-test: wast/wat_text_abbreviations/test_inline_export_on_func
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 42 call $log))
