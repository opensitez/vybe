;; vybe-test: wast/wat_execution/local_set_get_roundtrip
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    i32.const 99
    local.set $x
    local.get $x
    call $log))
