;; vybe-test: wast/wat_simd_i16x8/test_i16x8_sub_sat_s
;; origin: languages/wast/tests/wast/test_wat_simd_i16x8.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 v128.const i16x8 1 0 0 0 0 0 0 0
        i16x8.sub_sat_s i16x8.extract_lane_s 0 call $log)
)
