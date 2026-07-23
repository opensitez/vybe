//! Extended constant expressions proposal — global initializers may use
//! i32/i64 add/sub/mul and global.get of earlier immutable globals.
use crate::wat_exec;

wat_exec! {
    test_const_add_in_global_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.add (i32.const 40) (i32.const 2)))
        (func (export "_start") global.get $g call $log))"#, "42" },
    test_const_mul_in_global_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.mul (i32.const 6) (i32.const 7)))
        (func (export "_start") global.get $g call $log))"#, "42" },
    test_const_sub_in_global_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.sub (i32.const 100) (i32.const 58)))
        (func (export "_start") global.get $g call $log))"#, "42" },
    test_global_get_in_const_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $base i32 (i32.const 20))
        (global $derived i32 (i32.add (global.get $base) (i32.const 22)))
        (func (export "_start") global.get $derived call $log))"#, "42" },
    test_nested_const_expr => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.add (i32.mul (i32.const 5) (i32.const 8)) (i32.const 2)))
        (func (export "_start") global.get $g call $log))"#, "42" },
    test_i64_const_expr_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (global $g i64 (i64.mul (i64.const 1000000) (i64.const 1000000)))
        (func (export "_start") global.get $g call $log_i64))"#, "1000000000000" },
    test_const_expr_data_offset => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (global $off i32 (i32.const 8))
        (data (offset (i32.add (global.get $off) (i32.const 4))) "\63\00\00\00")
        (func (export "_start") i32.const 12 i32.load call $log))"#, "99" },
}
