;; vybe-test: wast/wat_simd_load_store/test_v128_load64_splat
;; origin: languages/wast/tests/wast/test_wat_simd_load_store.rs

(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (func $vybe_check_i64 (param i64) (param i64)
    local.get 0
    local.get 1
    i64.ne
    if
      unreachable
    end)
        (memory 1) (data (i32.const 0) "\01\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load64_splat i64x2.extract_lane 1 i64.const 1 call $vybe_check_i64))
