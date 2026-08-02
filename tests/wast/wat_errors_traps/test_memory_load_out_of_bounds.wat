;; vybe-test: wast/wat_errors_traps/test_memory_load_out_of_bounds
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs
;; vybe-test-mode: run-fail

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start") i32.const 100000 i32.load call $log))
