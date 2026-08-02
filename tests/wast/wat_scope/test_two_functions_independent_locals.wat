;; vybe-test: wast/wat_scope/test_two_functions_independent_locals
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $a (result i32) (local $x i32) i32.const 1 local.set $x local.get $x)
        (func $b (result i32) (local $x i32) i32.const 2 local.set $x local.get $x)
        (func (export "_start") call $a call $b i32.add call $log))
