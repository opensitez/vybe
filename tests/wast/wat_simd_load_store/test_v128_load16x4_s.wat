;; vybe-test: wast/wat_simd_load_store/test_v128_load16x4_s
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\ff\ff\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load16x4_s i32x4.extract_lane 0 call $log))
