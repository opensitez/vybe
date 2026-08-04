;; vybe-test: wast/wat_simd_i16x8/test_i16x8_shr_u
;; origin: languages/wast/tests/wast/test_wat_simd_i16x8.rs

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
        v128.const i16x8 0x8000 0 0 0 0 0 0 0 i32.const 1 i16x8.shr_u
        i16x8.extract_lane_u 0 i32.const 16384 call $vybe_check_i32)
)
