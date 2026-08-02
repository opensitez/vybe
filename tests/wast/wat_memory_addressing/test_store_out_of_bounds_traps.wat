;; vybe-test: wast/wat_memory_addressing/test_store_out_of_bounds_traps
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs
;; vybe-test-mode: run-fail

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 65536 i32.const 1 i32.store i32.const 0 call $log))
