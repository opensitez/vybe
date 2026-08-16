;; vybe-test: wast/wat_simd_load_store/test_v128_load16_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory 1) (data (i32.const 0) "\d2\04")
        (func (export "_start")
          ;; `v128.load16_lane : [i32 v128] -> [v128]` — ADDRESS first.
          i32.const 0
          v128.const i16x8 0 0 0 0 0 0 0 0
          v128.load16_lane 2 i16x8.extract_lane_u 2 i32.const 1234 call $vybe_check_i32))
