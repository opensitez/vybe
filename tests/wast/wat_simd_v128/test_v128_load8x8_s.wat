;; vybe-test: wast/wat_simd_v128/test_v128_load8x8_s
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\ff\00\80\7f\01\02\03\04")
(func (export "_start")
  i32.const 0
  v128.load8x8_s
  i16x8.extract_lane_s 2
  call $log
)
)
