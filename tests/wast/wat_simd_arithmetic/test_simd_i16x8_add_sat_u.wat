;; vybe-test: wast/wat_simd_arithmetic/test_simd_i16x8_add_sat_u
;; origin: languages/wast/tests/wast/test_wat_simd_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  v128.const i16x8 -1 0 0 0 0 0 0 0 ;; 65535
  v128.const i16x8 10 0 0 0 0 0 0 0
  i16x8.add_sat_u
  i16x8.extract_lane_u 0
  call $log
)
)
