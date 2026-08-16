;; vybe-test: wast/wat_simd_load_store/test_v128_load64_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (memory 1) (data (i32.const 0) "\07\00\00\00\00\00\00\00")
        (func (export "_start")
          ;; `v128.load64_lane : [i32 v128] -> [v128]` — ADDRESS first.
          i32.const 0
          v128.const i64x2 0 0
          v128.load64_lane 1 i64x2.extract_lane 1 i64.const 7 call $vybe_check_i64))
