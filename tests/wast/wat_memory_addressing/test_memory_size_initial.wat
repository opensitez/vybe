;; vybe-test: wast/wat_memory_addressing/test_memory_size_initial
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory 2)
        (func (export "_start") memory.size i32.const 2 call $vybe_check_i32))
