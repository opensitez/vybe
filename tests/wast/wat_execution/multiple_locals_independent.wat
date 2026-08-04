;; vybe-test: wast/wat_execution/multiple_locals_independent
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
    (local $a i32)
    (local $b i32)
    i32.const 10
    local.set $a
    i32.const 20
    local.set $b
    local.get $a
    i32.const 10 call $vybe_check_i32
    local.get $b
    i32.const 20 call $vybe_check_i32))
