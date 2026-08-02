;; vybe-test: wast/wat_simd_load_store/test_v128_load_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\63\00\00\00")
        (func (export "_start")
          v128.const i32x4 0 0 0 0 i32.const 0 v128.load32_lane 0 v128.load32_lane 1
          i32x4.extract_lane 1 call $log))
