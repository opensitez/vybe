;; vybe-test: wast/wat_extended_const/test_i64_const_expr_init
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (global $g i64 (i64.mul (i64.const 1000000) (i64.const 1000000)))
        (func (export "_start") global.get $g call $log_i64))
