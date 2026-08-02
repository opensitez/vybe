;; vybe-test: wast/wat_extended_const/test_nested_const_expr
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.add (i32.mul (i32.const 5) (i32.const 8)) (i32.const 2)))
        (func (export "_start") global.get $g call $log))
