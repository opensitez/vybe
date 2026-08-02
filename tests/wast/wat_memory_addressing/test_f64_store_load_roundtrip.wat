;; vybe-test: wast/wat_memory_addressing/test_f64_store_load_roundtrip
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 f64.const 3.14159 f64.store
          i32.const 0 f64.load call $log_f64))
