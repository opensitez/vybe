;; vybe-test: wast/wat_simd_load_store/test_v128_store16_lane
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i16x8 5000 0 0 0 0 0 0 0 v128.store16_lane 0
          i32.const 0 i32.load16_u i32.const 5000 call $vybe_check_i32))
