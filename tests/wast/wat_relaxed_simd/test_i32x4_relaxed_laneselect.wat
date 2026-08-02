;; vybe-test: wast/wat_relaxed_simd/test_i32x4_relaxed_laneselect
;; origin: languages/wast/tests/wast/test_wat_relaxed_simd.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i32x4 1 1 1 1 v128.const i32x4 2 2 2 2 v128.const i32x4 -1 0 0 0
        i32x4.relaxed_laneselect i32x4.extract_lane 0 call $log)
)
