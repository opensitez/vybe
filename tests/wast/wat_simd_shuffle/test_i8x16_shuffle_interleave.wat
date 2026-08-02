;; vybe-test: wast/wat_simd_shuffle/test_i8x16_shuffle_interleave
;; origin: languages/wast/tests/wast/test_wat_simd_shuffle.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9
        i8x16.shuffle 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23
        i8x16.extract_lane_u 1 call $log)
)
