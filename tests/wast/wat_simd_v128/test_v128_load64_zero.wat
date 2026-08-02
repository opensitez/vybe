;; vybe-test: wast/wat_simd_v128/test_v128_load64_zero
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\ff\ff\ff\ff\ff\ff\ff\ff")
(func (export "_start")
  i32.const 0
  v128.load64_zero
  i64x2.extract_lane 1
  call $log_i64
)
)
