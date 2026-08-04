;; vybe-test: wast/wat_simd_f32x4/test_f32x4_max_nan_propagates
;; origin: languages/wast/tests/wast/test_wat_simd_f32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_nan_f32 (param f32)
    local.get 0
    local.get 0
    f32.eq
    if
      unreachable
    end)
  (func (export "_start")
        v128.const f32x4 5.0 0 0 0 v128.const f32x4 nan 0 0 0
        f32x4.max f32x4.extract_lane 0 call $vybe_check_nan_f32)
)
