;; vybe-test: wast/wat_multi_memory/test_memory_size_of_second
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
        (memory $a 1) (memory $b 3)
        (func (export "_start") memory.size 1 i32.const 3 call $vybe_check_i32))
