;; vybe-test: wast/wat_scope/test_locals_are_per_call
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
        (func $f (param $p i32) (result i32) (local $x i32)
          local.get $p i32.const 1 i32.add local.set $x local.get $x)
        (func (export "_start")
          i32.const 10 call $f drop i32.const 20 call $f i32.const 21 call $vybe_check_i32))
