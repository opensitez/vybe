;; vybe-test: wast/wat_execution/memory_store_load_i8
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
  (memory 1)
  (func (export "_start")
    i32.const 4
    i32.const 200
    i32.store8
    i32.const 4
    i32.load8_u
    i32.const 200 call $vybe_check_i32))
