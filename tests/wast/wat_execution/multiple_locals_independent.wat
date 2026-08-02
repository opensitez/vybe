;; vybe-test: wast/wat_execution/multiple_locals_independent
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $a i32)
    (local $b i32)
    i32.const 10
    local.set $a
    i32.const 20
    local.set $b
    local.get $a
    call $log
    local.get $b
    call $log))
