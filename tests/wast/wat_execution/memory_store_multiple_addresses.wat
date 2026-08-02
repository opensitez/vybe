;; vybe-test: wast/wat_execution/memory_store_multiple_addresses
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0  i32.const 1 i32.store
    i32.const 4  i32.const 2 i32.store
    i32.const 8  i32.const 3 i32.store
    i32.const 0 i32.load call $log
    i32.const 4 i32.load call $log
    i32.const 8 i32.load call $log))
