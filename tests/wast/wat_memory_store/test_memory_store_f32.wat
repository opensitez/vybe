;; vybe-test: wast/wat_memory_store/test_memory_store_f32
;; origin: languages/wast/tests/wast/test_wat_memory_store.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f32 (param f32) (param f32)
    local.get 0
    local.get 1
    f32.ne
    if
      unreachable
    end)
  (memory 1)
(func (export "_start") 
  i32.const 0 
  f32.const 1.0 
  f32.store 
  i32.const 0 
  f32.load 
  f32.const 1.0 call $vybe_check_f32)
)
