;; vybe-test: wast/wat_simd_load_store/test_v128_load8_lane
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
        (memory 1) (data (i32.const 0) "\2a")
        (func (export "_start")
          ;; `v128.load8_lane : [i32 v128] -> [v128]` — ADDRESS first, then the
          ;; vector. This was written vector-first, which wasmtime rejects
          ;; outright; it only validated here because the emitter reversed the
          ;; two operands.
          i32.const 0
          v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
          v128.load8_lane 5 i8x16.extract_lane_u 5 i32.const 42 call $vybe_check_i32))
