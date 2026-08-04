;; vybe-test: wast/wat_memory64/test_i64_addressed_store_load
;; origin: languages/wast/tests/wast/test_wat_memory64.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory i64 1)
        (func (export "_start")
          i64.const 8 i32.const 42 i32.store
          i64.const 8 i32.load i32.const 42 call $vybe_check_i32))
