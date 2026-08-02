;; vybe-test: wast/wat_extended_const/test_const_expr_data_offset
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (global $off i32 (i32.const 8))
        (data (offset (i32.add (global.get $off) (i32.const 4))) "\63\00\00\00")
        (func (export "_start") i32.const 12 i32.load call $log))
