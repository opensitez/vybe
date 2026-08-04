;; vybe-test: wast/wat_simd_i32x4/test_i32x4_extmul_low_i16x8_s
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
        v128.const i16x8 1000 0 0 0 0 0 0 0 v128.const i16x8 1000 0 0 0 0 0 0 0
        i32x4.extmul_low_i16x8_s i32x4.extract_lane 0 i32.const 1000000 call $vybe_check_i32)
)
