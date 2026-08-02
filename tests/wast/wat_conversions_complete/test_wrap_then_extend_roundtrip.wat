;; vybe-test: wast/wat_conversions_complete/test_wrap_then_extend_roundtrip
;; origin: languages/wast/tests/wast/test_wat_conversions_complete.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i64.const 0x1_0000_002A i32.wrap_i64 i64.extend_i32_s call $log_i64)
)
