;; vybe-test: wast/wat_simd_load_store/test_v128_load_lane
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
        (memory 1) (data (i32.const 0) "\63\00\00\00")
        (func (export "_start") (local $v v128)
          ;; `v128.load32_lane : [i32 v128] -> [v128]` — ADDRESS first, so two
          ;; of them cannot be chained on the stack: the second needs its own
          ;; address UNDER the vector the first produced.
          i32.const 0
          v128.const i32x4 0 0 0 0
          v128.load32_lane 0
          local.set $v
          i32.const 0
          local.get $v
          v128.load32_lane 1
          i32x4.extract_lane 1 i32.const 99 call $vybe_check_i32))
