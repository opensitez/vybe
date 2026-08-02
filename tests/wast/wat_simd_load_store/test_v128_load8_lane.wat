;; vybe-test: wast/wat_simd_load_store/test_v128_load8_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a")
        (func (export "_start")
          v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
          i32.const 0 v128.load8_lane 5 i8x16.extract_lane_u 5 call $log))
