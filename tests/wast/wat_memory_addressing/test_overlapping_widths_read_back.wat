;; vybe-test: wast/wat_memory_addressing/test_overlapping_widths_read_back
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
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 0xAABBCCDD i32.store
          i32.const 2 i32.load16_u i32.const 43707 call $vybe_check_i32))
