;; vybe-test: wast/wat_simd_load_store/test_v128_load16x4_s
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
        (memory 1) (data (i32.const 0) "\ff\ff\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load16x4_s i32x4.extract_lane 0 i32.const -1 call $vybe_check_i32))
