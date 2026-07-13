//! Assignment concepts — writing to locals, globals, and memory, plus
//! compound updates, swaps, and read-modify-write patterns.
use crate::wat_exec;

wat_exec! {
    test_local_assignment => { r#"(func (export "_start")
        (local $x i32) i32.const 42 local.set $x local.get $x call $log)"#, "42" },
    test_reassignment_overwrites => { r#"(func (export "_start")
        (local $x i32) i32.const 1 local.set $x i32.const 2 local.set $x
        local.get $x call $log)"#, "2" },
    test_compound_increment => { r#"(func (export "_start")
        (local $x i32) i32.const 10 local.set $x
        local.get $x i32.const 5 i32.add local.set $x local.get $x call $log)"#, "15" },
    test_swap_two_locals => { r#"(func (export "_start")
        (local $a i32) (local $b i32) (local $t i32)
        i32.const 3 local.set $a i32.const 9 local.set $b
        local.get $a local.set $t local.get $b local.set $a local.get $t local.set $b
        local.get $a call $log)"#, "9" },
    test_swap_via_xor => { r#"(func (export "_start")
        (local $a i32) (local $b i32) i32.const 5 local.set $a i32.const 12 local.set $b
        local.get $a local.get $b i32.xor local.set $a
        local.get $a local.get $b i32.xor local.set $b
        local.get $a local.get $b i32.xor local.set $a
        local.get $a call $log)"#, "12" },
    test_global_assignment => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func (export "_start") i32.const 88 global.set $g global.get $g call $log))"#, "88" },
    test_global_compound_update => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 100))
        (func (export "_start")
          global.get $g i32.const 50 i32.sub global.set $g global.get $g call $log))"#, "50" },
    test_memory_assignment => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 1234 i32.store i32.const 0 i32.load call $log))"#, "1234" },
    test_memory_read_modify_write => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start")
          i32.const 0 i32.const 10 i32.store
          i32.const 0 i32.const 0 i32.load i32.const 5 i32.add i32.store
          i32.const 0 i32.load call $log))"#, "15" },
    test_multiple_assignments_evaluated_in_order => { r#"(func (export "_start")
        (local $a i32) (local $b i32) (local $c i32)
        i32.const 1 local.set $a
        local.get $a i32.const 1 i32.add local.set $b
        local.get $b i32.const 1 i32.add local.set $c
        local.get $c call $log)"#, "3" },
    test_tee_assigns_and_yields => { r#"(func (export "_start")
        (local $x i32) i32.const 7 local.tee $x local.get $x i32.add call $log)"#, "14" },
    test_conditional_assignment => { r#"(func (export "_start")
        (local $x i32) i32.const 1
        if i32.const 100 local.set $x else i32.const 200 local.set $x end
        local.get $x call $log)"#, "100" },
    test_accumulate_into_memory_cell => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func $add (param $n i32) i32.const 0 i32.const 0 i32.load local.get $n i32.add i32.store)
        (func (export "_start")
          i32.const 3 call $add i32.const 4 call $add i32.const 5 call $add
          i32.const 0 i32.load call $log))"#, "12" },
}
