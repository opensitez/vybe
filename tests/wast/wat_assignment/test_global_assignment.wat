;; vybe-test: wast/wat_assignment/test_global_assignment
;; origin: languages/wast/tests/wast/test_wat_assignment.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func (export "_start") i32.const 88 global.set $g global.get $g call $log))
