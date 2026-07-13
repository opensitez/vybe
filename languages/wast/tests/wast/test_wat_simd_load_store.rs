//! SIMD memory instructions — splat loads, lane loads/stores, widening loads,
//! and zero-extending loads, per the WebAssembly SIMD spec.
use crate::wat_exec;

wat_exec! {
    test_v128_load8_splat => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a")
        (func (export "_start")
          i32.const 0 v128.load8_splat i8x16.extract_lane_u 7 call $log))"#, "42" },
    test_v128_load16_splat => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\d2\04")
        (func (export "_start")
          i32.const 0 v128.load16_splat i16x8.extract_lane_u 3 call $log))"#, "1234" },
    test_v128_load32_splat => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\04\03\02\01")
        (func (export "_start")
          i32.const 0 v128.load32_splat i32x4.extract_lane 2 call $log))"#, "16909060" },
    test_v128_load64_splat => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load64_splat i64x2.extract_lane 1 call $log_i64))"#, "1" },
    test_v128_load_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\63\00\00\00")
        (func (export "_start")
          v128.const i32x4 0 0 0 0 i32.const 0 v128.load32_lane 0 v128.load32_lane 1
          i32x4.extract_lane 1 call $log))"#, "99" },
    test_v128_store_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i32x4 55 0 0 0 v128.store32_lane 0
          i32.const 0 i32.load call $log))"#, "55" },
    test_v128_load8x8_s => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\ff\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load8x8_s i16x8.extract_lane_s 0 call $log))"#, "-1" },
    test_v128_load8x8_u => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\ff\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load8x8_u i16x8.extract_lane_u 0 call $log))"#, "255" },
    test_v128_load16x4_s => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\ff\ff\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load16x4_s i32x4.extract_lane 0 call $log))"#, "-1" },
    test_v128_load32x2_u => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\ff\ff\ff\ff\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load32x2_u i64x2.extract_lane 0 call $log_i64))"#, "4294967295" },
    test_v128_load32_zero => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\07\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load32_zero i32x4.extract_lane 1 call $log))"#, "0" },
    test_v128_load64_zero => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\09\00\00\00\00\00\00\00")
        (func (export "_start")
          i32.const 0 v128.load64_zero i64x2.extract_lane 0 call $log_i64))"#, "9" },

    // ── remaining lane load/store widths (8/16/64) ───────────────────────────
    test_v128_load8_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a")
        (func (export "_start")
          v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
          i32.const 0 v128.load8_lane 5 i8x16.extract_lane_u 5 call $log))"#, "42" },
    test_v128_load16_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\d2\04")
        (func (export "_start")
          v128.const i16x8 0 0 0 0 0 0 0 0
          i32.const 0 v128.load16_lane 2 i16x8.extract_lane_u 2 call $log))"#, "1234" },
    test_v128_load64_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1) (data (i32.const 0) "\07\00\00\00\00\00\00\00")
        (func (export "_start")
          v128.const i64x2 0 0
          i32.const 0 v128.load64_lane 1 i64x2.extract_lane 1 call $log_i64))"#, "7" },
    test_v128_store8_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i8x16 99 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 v128.store8_lane 0
          i32.const 0 i32.load8_u call $log))"#, "99" },
    test_v128_store16_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i16x8 5000 0 0 0 0 0 0 0 v128.store16_lane 0
          i32.const 0 i32.load16_u call $log))"#, "5000" },
    test_v128_store64_lane => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 v128.const i64x2 123456789 0 v128.store64_lane 0
          i32.const 0 i64.load call $log_i64))"#, "123456789" },
}
