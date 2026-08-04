;; vybe-test: wast/wat_memory_addressing/test_i64_store_load_full_width
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0x0102030405060708 i64.store
          i32.const 0 i64.load i64.const 72623859790382856 call $vybe_check_i64))
