;; vybe-test: wast/wat_multivalue/test_func_returns_two_values
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $pair (result i32 i32) i32.const 11 i32.const 22)
        (func (export "_start") call $pair call $log call $log))
