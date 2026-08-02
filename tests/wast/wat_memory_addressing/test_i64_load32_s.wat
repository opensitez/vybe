;; vybe-test: wast/wat_memory_addressing/test_i64_load32_s
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0xFFFFFFFF i64.store32
          i32.const 0 i64.load32_s call $log_i64))
