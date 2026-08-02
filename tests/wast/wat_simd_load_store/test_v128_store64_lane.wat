;; vybe-test: wast/wat_simd_load_store/test_v128_store64_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i64x2 123456789 0 v128.store64_lane 0
          i32.const 0 i64.load call $log_i64))
