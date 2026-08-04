;; vybe-test: wast/wat_multi_memory/test_first_memory_unaffected_by_second
;; origin: languages/wast/tests/wast/test_wat_multi_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory $a 1) (memory $b 1)
        (func (export "_start")
          i32.const 0 i32.const 111 i32.store 0
          i32.const 0 i32.const 222 i32.store 1
          i32.const 0 i32.load 0 i32.const 111 call $vybe_check_i32))
