;; vybe-test: wast/wat_execution/global_immutable_read
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
  (global $c i32 (i32.const 42))
  (func (export "_start")
    global.get $c
    i32.const 42 call $vybe_check_i32))
