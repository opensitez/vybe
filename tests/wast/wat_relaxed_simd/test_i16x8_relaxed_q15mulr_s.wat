;; vybe-test: wast/wat_relaxed_simd/test_i16x8_relaxed_q15mulr_s
;; origin: languages/wast/tests/wast/test_wat_relaxed_simd.rs

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
        v128.const i16x8 16384 0 0 0 0 0 0 0 v128.const i16x8 16384 0 0 0 0 0 0 0
        i16x8.relaxed_q15mulr_s i16x8.extract_lane_s 0 i32.const 8192 call $vybe_check_i32)
)
