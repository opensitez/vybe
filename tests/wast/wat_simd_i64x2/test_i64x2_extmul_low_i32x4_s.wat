;; vybe-test: wast/wat_simd_i64x2/test_i64x2_extmul_low_i32x4_s
;; origin: languages/wast/tests/wast/test_wat_simd_i64x2.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
  (func (export "_start")
        v128.const i32x4 100000 0 0 0 v128.const i32x4 100000 0 0 0
        i64x2.extmul_low_i32x4_s i64x2.extract_lane 0 i64.const 10000000000 call $vybe_check_i64)
)
