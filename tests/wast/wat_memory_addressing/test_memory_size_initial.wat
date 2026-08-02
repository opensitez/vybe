;; vybe-test: wast/wat_memory_addressing/test_memory_size_initial
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 2)
        (func (export "_start") memory.size call $log))
