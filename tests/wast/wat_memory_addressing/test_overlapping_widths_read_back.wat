;; vybe-test: wast/wat_memory_addressing/test_overlapping_widths_read_back
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 0xAABBCCDD i32.store
          i32.const 2 i32.load16_u call $log))
