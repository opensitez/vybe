;; vybe-test: wast/wat_execution_extended/test_memory_load_alignments
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0
    i32.const 99
    i32.store align=4
    i32.const 0
    i32.load align=1
    call $log))
