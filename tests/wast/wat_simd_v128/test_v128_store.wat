;; vybe-test: wast/wat_simd_v128/test_v128_store
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start")
  i32.const 0
  v128.const i32x4 42 99 100 200
  v128.store
  i32.const 4
  i32.load
  call $log
)
)
