;; vybe-test: wast/wat_memory64/test_i64_memory_size
;; origin: languages/wast/tests/wast/test_wat_memory64.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory i64 2)
        (func (export "_start") memory.size call $log_i64))
