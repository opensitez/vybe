;; vybe-test: wast/wat_simd_v128/test_v128_const
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

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
  v128.const i32x4 0x01020304 0x05060708 0x090A0B0C 0x0D0E0F10
  i32x4.extract_lane 0
  i32.const 16909060 call $vybe_check_i32
)
)
