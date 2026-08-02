;; vybe-test: wast/wat_simd_i8x16/test_i8x16_sub_sat_s_clamps_low
;; origin: languages/wast/tests/wast/test_wat_simd_i8x16.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.sub_sat_s i8x16.extract_lane_s 0 call $log)
)
