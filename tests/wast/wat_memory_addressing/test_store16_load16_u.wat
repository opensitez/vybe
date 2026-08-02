;; vybe-test: wast/wat_memory_addressing/test_store16_load16_u
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 4 i32.const 40000 i32.store16
          i32.const 4 i32.load16_u call $log))
