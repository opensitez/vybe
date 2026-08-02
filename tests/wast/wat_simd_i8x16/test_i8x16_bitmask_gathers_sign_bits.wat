;; vybe-test: wast/wat_simd_i8x16/test_i8x16_bitmask_gathers_sign_bits
;; origin: languages/wast/tests/wast/test_wat_simd_i8x16.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 0x80 0 0x80 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.bitmask call $log)
)
