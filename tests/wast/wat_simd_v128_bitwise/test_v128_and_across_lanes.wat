;; vybe-test: wast/wat_simd_v128_bitwise/test_v128_and_across_lanes
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
        v128.const i16x8 0xFFFF 0xFFFF 0 0 0 0 0 0
        v128.const i16x8 0x00FF 0xFF00 0 0 0 0 0 0
        v128.and i16x8.extract_lane_u 1 i32.const 65280 call $vybe_check_i32)
)
