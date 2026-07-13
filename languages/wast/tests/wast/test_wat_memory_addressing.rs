//! Memory addressing: static offsets, load/store width and sign, and
//! out-of-bounds traps. Little-endian layout per the WebAssembly spec.
use crate::wat_exec;

wat_exec! {
    test_load_with_static_offset => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 12345 i32.store offset=8
          i32.const 0 i32.load offset=8 call $log))"#, "12345" },
    test_store_load_i8_zero_extends => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 200 i32.store8
          i32.const 0 i32.load8_u call $log))"#, "200" },
    test_load8_s_sign_extends => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 200 i32.store8
          i32.const 0 i32.load8_s call $log))"#, "-56" },
    test_store16_load16_u => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 4 i32.const 40000 i32.store16
          i32.const 4 i32.load16_u call $log))"#, "40000" },
    test_load16_s_sign_extends => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 4 i32.const 0xFFFF i32.store16
          i32.const 4 i32.load16_s call $log))"#, "-1" },
    test_i64_store_load_full_width => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0x0102030405060708 i64.store
          i32.const 0 i64.load call $log_i64))"#, "72623859790382856" },
    test_i64_load32_u => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0xFFFFFFFF i64.store32
          i32.const 0 i64.load32_u call $log_i64))"#, "4294967295" },
    test_i64_load32_s => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i64.const 0xFFFFFFFF i64.store32
          i32.const 0 i64.load32_s call $log_i64))"#, "-1" },
    test_f64_store_load_roundtrip => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
        (memory 1)
        (func (export "_start")
          i32.const 0 f64.const 3.14159 f64.store
          i32.const 0 f64.load call $log_f64))"#, "3.14159" },
    test_little_endian_byte_order => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 0x04030201 i32.store
          i32.const 0 i32.load8_u call $log))"#, "1" },
    test_load_out_of_bounds_traps => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 70000 i32.load call $log))"#, "trap" },
    test_store_out_of_bounds_traps => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 65536 i32.const 1 i32.store i32.const 0 call $log))"#, "trap" },
    test_memory_size_initial => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 2)
        (func (export "_start") memory.size call $log))"#, "2" },
    test_overlapping_widths_read_back => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 0xAABBCCDD i32.store
          i32.const 2 i32.load16_u call $log))"#, "43707" },
}
