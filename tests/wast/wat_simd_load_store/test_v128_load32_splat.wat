;; vybe-test: wast/wat_simd_load_store/test_v128_load32_splat
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\04\03\02\01")
        (func (export "_start")
          i32.const 0 v128.load32_splat i32x4.extract_lane 2 call $log))
