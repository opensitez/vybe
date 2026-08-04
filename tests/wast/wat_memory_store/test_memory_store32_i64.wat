;; vybe-test: wast/wat_memory_store/test_memory_store32_i64
;; origin: languages/wast/tests/wast/test_wat_memory_store.rs

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
(func (export "_start") 
  i32.const 0 
  i64.const 4294967340 ;; 4294967296 + 44
  i64.store32 
  i32.const 0 
  i64.load32_u 
  i64.const 44 call $vybe_check_i64)
)
