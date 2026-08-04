;; vybe-test: wast/wat_memory_store/test_memory_store16_i32
;; origin: languages/wast/tests/wast/test_wat_memory_store.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (memory 1)
(func (export "_start") 
  i32.const 0 
  i32.const 65580 ;; 65536 + 44
  i32.store16 
  i32.const 0 
  i32.load16_u 
  i32.const 44 call $vybe_check_i32)
)
