;; vybe-test: wast/wat_memory_addressing/test_load_with_static_offset
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 12345 i32.store offset=8
          i32.const 0 i32.load offset=8 call $log))
