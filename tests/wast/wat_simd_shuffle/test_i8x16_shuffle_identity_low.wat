;; vybe-test: wast/wat_simd_shuffle/test_i8x16_shuffle_identity_low
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
        v128.const i8x16 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25
        v128.const i8x16 100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115
        i8x16.shuffle 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15
        i8x16.extract_lane_u 3 i32.const 13 call $vybe_check_i32)
)
