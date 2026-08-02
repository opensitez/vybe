;; vybe-test: wast/wat_memory_addressing/test_i64_store_load_full_width
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0x0102030405060708 i64.store
          i32.const 0 i64.load call $log_i64))
