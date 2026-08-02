;; vybe-test: wast/wat_memory64/test_i64_address_high_offset
;; origin: languages/wast/tests/wast/test_wat_memory64.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory i64 1)
        (func (export "_start")
          i64.const 1000 i32.const 777 i32.store
          i64.const 1000 i32.load call $log))
