;; vybe-test: wast/wat_execution/direct_call_add
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
  (func $add (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.add)
  (func (export "_start")
    i32.const 13
    i32.const 29
    call $add
    i32.const 42 call $vybe_check_i32))
