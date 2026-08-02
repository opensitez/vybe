;; vybe-test: wast/wat_simd_load_store/test_v128_store16_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i16x8 5000 0 0 0 0 0 0 0 v128.store16_lane 0
          i32.const 0 i32.load16_u call $log))
