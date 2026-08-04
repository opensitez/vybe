;; vybe-test: wast/wat_simd_f64x2/test_f64x2_sqrt
;; origin: languages/wast/tests/wast/test_wat_simd_f64x2.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f64 (param f64) (param f64)
    local.get 0
    local.get 1
    f64.ne
    if
      unreachable
    end)
  (func (export "_start")
        v128.const f64x2 25.0 0 f64x2.sqrt f64x2.extract_lane 0 f64.const 5.0 call $vybe_check_f64)
)
