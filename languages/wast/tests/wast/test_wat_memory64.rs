//! Memory64 proposal — a memory declared with an i64 index type; addresses are
//! 64-bit values.
use crate::wat_exec;

wat_exec! {
    test_i64_addressed_store_load => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory i64 1)
        (func (export "_start")
          i64.const 8 i32.const 42 i32.store
          i64.const 8 i32.load call $log))"#, "42" },
    test_i64_address_high_offset => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory i64 1)
        (func (export "_start")
          i64.const 1000 i32.const 777 i32.store
          i64.const 1000 i32.load call $log))"#, "777" },
    test_i64_memory_size => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (memory i64 2)
        (func (export "_start") memory.size call $log_i64))"#, "2" },
    test_i64_addressed_byte => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory i64 1)
        (func (export "_start")
          i64.const 0 i32.const 200 i32.store8
          i64.const 0 i32.load8_u call $log))"#, "200" },
}
