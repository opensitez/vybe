;; vybe-test: wast/wat_simd_shuffle/test_i8x16_swizzle_by_index
;; origin: languages/wast/tests/wast/test_wat_simd_shuffle.rs

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
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 130 140 150 160
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.swizzle i8x16.extract_lane_u 0 i32.const 60 call $vybe_check_i32)
)
