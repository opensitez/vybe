;; vybe-test: wast/wat_execution/memory_store_load_i32
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
    i32.const 0     ;; address
    i32.const 12345
    i32.store
    i32.const 0
    i32.load
    i32.const 12345 call $vybe_check_i32))
