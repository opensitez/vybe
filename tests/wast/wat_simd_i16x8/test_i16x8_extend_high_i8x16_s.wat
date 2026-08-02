;; vybe-test: wast/wat_simd_i16x8/test_i16x8_extend_high_i8x16_s
;; origin: languages/wast/tests/wast/test_wat_simd_i16x8.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0x80 0 0 0 0 0 0 0
        i16x8.extend_high_i8x16_s i16x8.extract_lane_s 0 call $log)
)
