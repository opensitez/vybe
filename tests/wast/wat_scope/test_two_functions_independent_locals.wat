;; vybe-test: wast/wat_scope/test_two_functions_independent_locals
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
        (func $a (result i32) (local $x i32) i32.const 1 local.set $x local.get $x)
        (func $b (result i32) (local $x i32) i32.const 2 local.set $x local.get $x)
        (func (export "_start") call $a call $b i32.add i32.const 3 call $vybe_check_i32))
