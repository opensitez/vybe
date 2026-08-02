;; vybe-test: wast/wat_simd_load_store/test_v128_store8_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i8x16 99 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 v128.store8_lane 0
          i32.const 0 i32.load8_u call $log))
