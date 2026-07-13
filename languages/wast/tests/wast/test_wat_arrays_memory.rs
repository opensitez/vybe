//! Array concepts implemented over linear memory — indexing, iteration, sum,
//! max, search, reverse, and in-place update.
use crate::wat_exec;

wat_exec! {
    test_array_sum => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\02\00\00\00\03\00\00\00\04\00\00\00")
        (func (export "_start") (local $i i32) (local $s i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $s local.get $i i32.load i32.add local.set $s
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          local.get $s call $log))"#, "10" },
    test_array_max => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\03\00\00\00\09\00\00\00\05\00\00\00\07\00\00\00")
        (func (export "_start") (local $i i32) (local $m i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $i i32.load local.get $m i32.gt_s
            if local.get $i i32.load local.set $m end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          local.get $m call $log))"#, "9" },
    test_array_linear_search_found => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\0a\00\00\00\14\00\00\00\1e\00\00\00\28\00\00\00")
        (func (export "_start") (local $i i32)
          block loop local.get $i i32.const 16 i32.ge_u
            if i32.const -1 call $log return end
            local.get $i i32.load i32.const 30 i32.eq
            if local.get $i i32.const 4 i32.div_u call $log return end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end))"#, "2" },
    test_array_write_then_read => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1)
        (func (export "_start") (local $i i32)
          block loop local.get $i i32.const 10 i32.ge_u br_if 1
            local.get $i i32.const 4 i32.mul local.get $i local.get $i i32.mul i32.store
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          i32.const 28 i32.load call $log))"#, "49" },
    test_array_reverse_in_place => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\02\00\00\00\03\00\00\00\04\00\00\00")
        (func (export "_start") (local $l i32) (local $r i32) (local $t i32)
          i32.const 12 local.set $r
          block loop local.get $l local.get $r i32.ge_u br_if 1
            local.get $l i32.load local.set $t
            local.get $l local.get $r i32.load i32.store
            local.get $r local.get $t i32.store
            local.get $l i32.const 4 i32.add local.set $l
            local.get $r i32.const 4 i32.sub local.set $r br 0 end end
          i32.const 0 i32.load call $log))"#, "4" },
    test_array_count_matching => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\02\00\00\00\04\00\00\00\05\00\00\00\06\00\00\00")
        (func (export "_start") (local $i i32) (local $c i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $i i32.load i32.const 1 i32.and i32.eqz
            if local.get $c i32.const 1 i32.add local.set $c end
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          local.get $c call $log))"#, "3" },
    test_byte_array_dot_product => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\02\03") (data (i32.const 8) "\04\05\06")
        (func (export "_start") (local $i i32) (local $s i32)
          block loop local.get $i i32.const 3 i32.ge_u br_if 1
            local.get $s local.get $i i32.load8_u local.get $i i32.const 8 i32.add i32.load8_u i32.mul i32.add local.set $s
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          local.get $s call $log))"#, "32" },
    test_prefix_sum_last => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\01\00\00\00\02\00\00\00\03\00\00\00\04\00\00\00")
        (func (export "_start") (local $i i32) (local $run i32)
          block loop local.get $i i32.const 16 i32.ge_u br_if 1
            local.get $run local.get $i i32.load i32.add local.set $run
            local.get $i local.get $run i32.store
            local.get $i i32.const 4 i32.add local.set $i br 0 end end
          i32.const 12 i32.load call $log))"#, "10" },
}
