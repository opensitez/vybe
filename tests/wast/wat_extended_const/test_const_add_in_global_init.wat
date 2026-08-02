;; vybe-test: wast/wat_extended_const/test_const_add_in_global_init
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.add (i32.const 40) (i32.const 2)))
        (func (export "_start") global.get $g call $log))
