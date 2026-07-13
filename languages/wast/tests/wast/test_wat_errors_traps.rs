//! Error / trap concepts — every way a WebAssembly program can trap, and the
//! guard patterns that avoid them. Traps use the `"trap"` expectation.
use crate::wat_exec;

wat_exec! {
    test_integer_divide_by_zero => { r#"(func (export "_start")
        i32.const 10 i32.const 0 i32.div_s call $log)"#, "trap" },
    test_unsigned_divide_by_zero => { r#"(func (export "_start")
        i32.const 10 i32.const 0 i32.div_u call $log)"#, "trap" },
    test_remainder_by_zero => { r#"(func (export "_start")
        i32.const 10 i32.const 0 i32.rem_s call $log)"#, "trap" },
    test_i64_divide_by_zero => { r#"(func (export "_start")
        i64.const 10 i64.const 0 i64.div_s call $log_i64)"#, "trap" },
    test_signed_overflow_div => { r#"(func (export "_start")
        i32.const -2147483648 i32.const -1 i32.div_s call $log)"#, "trap" },
    test_explicit_unreachable => { r#"(func (export "_start")
        unreachable)"#, "trap" },
    test_float_trunc_out_of_range => { r#"(func (export "_start")
        f32.const 1e30 i32.trunc_f32_s call $log)"#, "trap" },
    test_float_trunc_nan => { r#"(func (export "_start")
        f64.const nan i32.trunc_f64_u call $log)"#, "trap" },
    test_memory_load_out_of_bounds => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start") i32.const 100000 i32.load call $log))"#, "trap" },
    test_memory_store_out_of_bounds => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start") i32.const 100000 i32.const 1 i32.store i32.const 0 call $log))"#, "trap" },
    test_conditional_trap_taken => { r#"(func (export "_start")
        i32.const 1 if unreachable end i32.const 0 call $log)"#, "trap" },
    test_conditional_trap_not_taken => { r#"(func (export "_start")
        i32.const 0 if unreachable end i32.const 99 call $log)"#, "99" },
    test_guard_avoids_divide_by_zero => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $safediv (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) i32.const -1 else local.get $a local.get $b i32.div_s end)
        (func (export "_start") i32.const 10 i32.const 0 call $safediv call $log))"#, "-1" },
    test_bounds_check_before_load => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func $safeload (param $addr i32) (result i32)
          local.get $addr i32.const 65536 i32.ge_u
          if (result i32) i32.const -1 else local.get $addr i32.load end)
        (func (export "_start") i32.const 999999 call $safeload call $log))"#, "-1" },
    test_trap_propagates_through_call => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $boom (result i32) unreachable)
        (func (export "_start") call $boom call $log))"#, "trap" },
}
