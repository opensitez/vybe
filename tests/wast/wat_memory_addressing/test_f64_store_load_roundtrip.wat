;; vybe-test: wast/wat_memory_addressing/test_f64_store_load_roundtrip
;; origin: languages/wast/tests/wast/test_wat_memory_addressing.rs

(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f64 (param f64) (param f64)
    local.get 0
    local.get 1
    f64.ne
    if
      unreachable
    end)
        (memory 1)
        (func (export "_start")
          i32.const 0 f64.const 3.14159 f64.store
          i32.const 0 f64.load f64.const 3.14159 call $vybe_check_f64))
