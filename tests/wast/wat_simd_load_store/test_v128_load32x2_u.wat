;; vybe-test: wast/wat_simd_load_store/test_v128_load32x2_u
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\ff\ff\ff\ff\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load32x2_u i64x2.extract_lane 0 call $log_i64))
