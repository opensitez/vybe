;; vybe-test: wast/wat_memory_addressing/test_store_load_i8_zero_extends
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 200 i32.store8
          i32.const 0 i32.load8_u call $log))
