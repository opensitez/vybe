;; vybe-test: wast/wat_simd_i32x4/test_i32x4_trunc_sat_f64x2_s_zero_upper_lane_is_zero
;; origin: languages/wast/tests/wast/test_wat_simd_i32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
        v128.const f64x2 3.9 7.9 i32x4.trunc_sat_f64x2_s_zero i32x4.extract_lane 2 i32.const 0 call $vybe_check_i32)
)
