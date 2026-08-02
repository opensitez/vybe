;; vybe-test: wast/wat_simd_load_store/test_v128_load16_splat
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\d2\04")
        (func (export "_start")
          i32.const 0 v128.load16_splat i16x8.extract_lane_u 3 call $log))
