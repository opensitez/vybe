;; vybe-test: wast/wat_execution/memory_store_load_i32
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0     ;; address
    i32.const 12345
    i32.store
    i32.const 0
    i32.load
    call $log))
