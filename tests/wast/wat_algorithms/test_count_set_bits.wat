;; vybe-test: wast/wat_algorithms/test_count_set_bits
;; origin: languages/wast/tests/wast/test_wat_algorithms.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 0xB7 i32.popcnt call $log))
