//! Global variables as shared state — initialization, mutation across calls,
//! and use as counters/accumulators.
use crate::wat_exec;

wat_exec! {
    test_global_immutable_read => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.const 314))
        (func (export "_start") global.get $g call $log))"#, "314" },
    test_global_mutable_set => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func (export "_start") i32.const 77 global.set $g global.get $g call $log))"#, "77" },
    test_global_counter_across_calls => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $c (mut i32) (i32.const 0))
        (func $inc global.get $c i32.const 1 i32.add global.set $c)
        (func (export "_start")
          call $inc call $inc call $inc call $inc global.get $c call $log))"#, "4" },
    test_global_f64_init => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
        (global $pi f64 (f64.const 3.5))
        (func (export "_start") global.get $pi call $log_f64))"#, "3.5" },
    test_global_i64_accumulator => { r#"(module
        (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
        (global $acc (mut i64) (i64.const 0))
        (func $add (param $n i64) global.get $acc local.get $n i64.add global.set $acc)
        (func (export "_start")
          i64.const 1000000000 call $add i64.const 2000000000 call $add
          global.get $acc call $log_i64))"#, "3000000000" },
    test_global_initialized_from_other_global => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $base i32 (i32.const 10))
        (global $derived i32 (global.get $base))
        (func (export "_start") global.get $derived call $log))"#, "10" },
    test_global_toggle_flag => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $flag (mut i32) (i32.const 0))
        (func $toggle global.get $flag i32.eqz global.set $flag)
        (func (export "_start") call $toggle call $toggle call $toggle global.get $flag call $log))"#, "1" },
    test_global_decrement_loop => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $n (mut i32) (i32.const 3))
        (func (export "_start")
          block loop
            global.get $n i32.eqz br_if 1
            global.get $n i32.const 1 i32.sub global.set $n br 0
          end end
          global.get $n call $log))"#, "0" },
}
