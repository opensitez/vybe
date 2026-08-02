;; vybe-test: wast/wat_simd_v128/test_v128_const
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  v128.const i32x4 0x01020304 0x05060708 0x090A0B0C 0x0D0E0F10
  i32x4.extract_lane 0
  call $log
)
)
