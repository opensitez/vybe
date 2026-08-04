;; vybe-test: wast/wat_simd_v128_bitwise/test_v128_bitselect
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
        v128.const i32x4 0xAAAA 0 0 0
        v128.const i32x4 0x5555 0 0 0
        v128.const i32x4 0xFF00 0 0 0
        v128.bitselect i32x4.extract_lane 0 i32.const 43605 call $vybe_check_i32)
)
