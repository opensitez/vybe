;; vybe-test: wast/wat_simd_load_store/test_v128_load32_zero
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
        (memory 1) (data (i32.const 0) "\07\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load32_zero i32x4.extract_lane 1 i32.const 0 call $vybe_check_i32))
