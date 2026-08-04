;; vybe-test: wast/wat_simd_f32x4/test_f32x4_neg
;; origin: languages/wast/tests/wast/test_wat_simd_f32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f32 (param f32) (param f32)
    local.get 0
    local.get 1
    f32.ne
    if
      unreachable
    end)
  (func (export "_start")
        v128.const f32x4 7.5 0 0 0 f32x4.neg f32x4.extract_lane 0 f32.const -7.5 call $vybe_check_f32)
)
