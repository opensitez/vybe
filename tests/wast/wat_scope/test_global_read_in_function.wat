;; vybe-test: wast/wat_scope/test_global_read_in_function
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $base i32 (i32.const 1000))
        (func $offset (param $d i32) (result i32) global.get $base local.get $d i32.add)
        (func (export "_start") i32.const 23 call $offset call $log))
