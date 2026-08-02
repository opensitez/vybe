;; vybe-test: wast/wat_simd_i16x8/test_i16x8_shl
;; origin: languages/wast/tests/wast/test_wat_simd_i16x8.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i16x8 1 0 0 0 0 0 0 0 i32.const 4 i16x8.shl i16x8.extract_lane_s 0 call $log)
)
