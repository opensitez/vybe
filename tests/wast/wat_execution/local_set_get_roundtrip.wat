;; vybe-test: wast/wat_execution/local_set_get_roundtrip
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
    (local $x i32)
    i32.const 99
    local.set $x
    local.get $x
    i32.const 99 call $vybe_check_i32))
