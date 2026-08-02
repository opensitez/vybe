;; vybe-test: wast/wat_scope/test_global_shared_across_calls
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func $bump global.get $g i32.const 1 i32.add global.set $g)
        (func (export "_start") call $bump call $bump call $bump global.get $g call $log))
