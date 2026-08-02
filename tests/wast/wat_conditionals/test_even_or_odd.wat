;; vybe-test: wast/wat_conditionals/test_even_or_odd
;; origin: languages/wast/tests/wast/test_wat_conditionals.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $parity (param $n i32) (result i32) local.get $n i32.const 1 i32.and)
        (func (export "_start") i32.const 13 call $parity call $log))
