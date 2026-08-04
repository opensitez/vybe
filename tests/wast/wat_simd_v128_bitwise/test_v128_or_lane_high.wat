;; vybe-test: wast/wat_simd_v128_bitwise/test_v128_or_lane_high
;; origin: languages/wast/tests/wast/test_wat_simd_v128_bitwise.rs

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
        v128.const i32x4 0 0 0 0x1 v128.const i32x4 0 0 0 0x2
        v128.or i32x4.extract_lane 3 i32.const 3 call $vybe_check_i32)
)
