;; vybe-test: wast/wat_scope/test_local_does_not_leak_to_caller
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $inner (result i32) (local $x i32) i32.const 99 local.set $x local.get $x)
        (func (export "_start") (local $x i32)
          i32.const 5 local.set $x call $inner drop local.get $x call $log))
