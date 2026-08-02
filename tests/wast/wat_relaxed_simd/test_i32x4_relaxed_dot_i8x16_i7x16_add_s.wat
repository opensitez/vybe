;; vybe-test: wast/wat_relaxed_simd/test_i32x4_relaxed_dot_i8x16_i7x16_add_s
;; origin: languages/wast/tests/wast/test_wat_relaxed_simd.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 1 2 3 4 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 2 2 2 2 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i32x4 100 0 0 0
        i32x4.relaxed_dot_i8x16_i7x16_add_s i32x4.extract_lane 0 call $log)
)
