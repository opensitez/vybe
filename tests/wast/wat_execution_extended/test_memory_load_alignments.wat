;; vybe-test: wast/wat_execution_extended/test_memory_load_alignments
;; origin: languages/wast/tests/wast/test_wat_execution_extended.rs

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
    i32.const 0
    i32.const 99
    i32.store align=4
    i32.const 0
    i32.load align=1
    i32.const 99 call $vybe_check_i32))
