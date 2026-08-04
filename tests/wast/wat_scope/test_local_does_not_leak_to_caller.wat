;; vybe-test: wast/wat_scope/test_local_does_not_leak_to_caller
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (func $inner (result i32) (local $x i32) i32.const 99 local.set $x local.get $x)
        (func (export "_start") (local $x i32)
          i32.const 5 local.set $x call $inner drop local.get $x i32.const 5 call $vybe_check_i32))
