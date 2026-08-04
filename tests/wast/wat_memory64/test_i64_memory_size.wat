;; vybe-test: wast/wat_memory64/test_i64_memory_size
;; origin: languages/wast/tests/wast/test_wat_memory64.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (memory i64 2)
        (func (export "_start") memory.size i64.const 2 call $vybe_check_i64))
