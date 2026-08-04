;; vybe-test: wast/wat_relaxed_simd/test_i8x16_relaxed_laneselect
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
        v128.const i8x16 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1
        v128.const i8x16 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
        v128.const i8x16 0xFF 0 0xFF 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.relaxed_laneselect i8x16.extract_lane_u 0 i32.const 1 call $vybe_check_i32)
)
