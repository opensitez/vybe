;; vybe-test: wast/wat_memory_load/test_memory_load_i64
;; origin: languages/wast/tests/wast/test_wat_memory_load.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
  (memory 1)
(data (i32.const 8) "\01\02\03\04\05\06\07\08")
(func (export "_start") i32.const 8 i64.load i64.const 578437695752307201 call $vybe_check_i64)
)
