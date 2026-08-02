;; vybe-test: wast/wat_multi_memory/test_memory_size_of_second
;; origin: languages/wast/tests/wast/test_wat_multi_memory.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $a 1) (memory $b 3)
        (func (export "_start") memory.size 1 call $log))
