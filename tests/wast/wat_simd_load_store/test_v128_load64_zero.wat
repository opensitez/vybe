;; vybe-test: wast/wat_simd_load_store/test_v128_load64_zero
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\09\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load64_zero i64x2.extract_lane 0 call $log_i64))
