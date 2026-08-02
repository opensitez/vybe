;; vybe-test: wast/wat_memory64/test_i64_addressed_byte
;; origin: languages/wast/tests/wast/test_wat_memory64.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory i64 1)
        (func (export "_start")
          i64.const 0 i32.const 200 i32.store8
          i64.const 0 i32.load8_u call $log))
