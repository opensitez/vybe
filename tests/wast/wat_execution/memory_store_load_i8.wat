;; vybe-test: wast/wat_execution/memory_store_load_i8
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 4
    i32.const 200
    i32.store8
    i32.const 4
    i32.load8_u
    call $log))
