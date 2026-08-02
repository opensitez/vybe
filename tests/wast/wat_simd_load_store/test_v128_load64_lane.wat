;; vybe-test: wast/wat_simd_load_store/test_v128_load64_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\07\00\00\00\00\00\00\00")
        (func (export "_start")
          v128.const i64x2 0 0
          i32.const 0 v128.load64_lane 1 i64x2.extract_lane 1 call $log_i64))
