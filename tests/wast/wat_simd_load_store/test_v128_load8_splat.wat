;; vybe-test: wast/wat_simd_load_store/test_v128_load8_splat
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a")
        (func (export "_start")
          i32.const 0 v128.load8_splat i8x16.extract_lane_u 7 call $log))
