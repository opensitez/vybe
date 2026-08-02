;; vybe-test: wast/wat_simd_v128/test_v128_load16x4_u
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\ff\ff\00\00\00\80\ff\7f")
(func (export "_start")
  i32.const 0
  v128.load16x4_u
  i32x4.extract_lane 2
  call $log
)
)
