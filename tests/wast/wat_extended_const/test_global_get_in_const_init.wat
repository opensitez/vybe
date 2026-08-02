;; vybe-test: wast/wat_extended_const/test_global_get_in_const_init
;; origin: languages/wast/tests/wast/test_wat_extended_const.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $base i32 (i32.const 20))
        (global $derived i32 (i32.add (global.get $base) (i32.const 22)))
        (func (export "_start") global.get $derived call $log))
