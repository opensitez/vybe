//! Deeply nested control flow, branch depth targeting, and structured jumps.
use crate::wat_exec;

wat_exec! {
    test_br_to_outer_block_by_depth => { r#"(func (export "_start")
        block block block i32.const 9 call $log br 2 i32.const 99 call $log end
        i32.const 88 call $log end end)"#, "9" },
    test_br_to_inner_block_only => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func (export "_start")
          block block i32.const 1 call $log br 0 i32.const 2 call $log end
          i32.const 3 call $log end))"#, "1" },
    test_br_if_true_takes_branch => { r#"(func (export "_start")
        block i32.const 7 call $log i32.const 1 br_if 0 i32.const 8 call $log end)"#, "7" },
    test_br_if_false_falls_through => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func (export "_start")
          block i32.const 1 call $log i32.const 0 br_if 0 i32.const 2 call $log end))"#, "1" },
    test_nested_loop_accumulate => { r#"(func (export "_start")
        (local $sum i32) (local $i i32)
        i32.const 5 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $sum local.get $i i32.add local.set $sum
          local.get $i i32.const 1 i32.sub local.set $i
          br 0
        end end
        local.get $sum call $log)"#, "15" },
    test_br_table_selects_middle => { r#"(func (export "_start")
        block block block block
          i32.const 1 br_table 0 1 2 3
        end i32.const 100 call $log br 2
        end i32.const 200 call $log br 1
        end i32.const 300 call $log br 0
        end)"#, "200" },
    test_br_table_default_target => { r#"(func (export "_start")
        block block
          i32.const 9 br_table 0 1
        end i32.const 111 call $log br 1
        end i32.const 222 call $log)"#, "222" },
    test_return_exits_function_early => { r#"(func (export "_start")
        i32.const 5 call $log return i32.const 6 call $log)"#, "5" },
    test_if_else_nested_in_loop => { r#"(func (export "_start")
        (local $i i32) (local $acc i32)
        i32.const 4 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $i i32.const 2 i32.rem_u i32.eqz
          if local.get $acc local.get $i i32.add local.set $acc end
          local.get $i i32.const 1 i32.sub local.set $i
          br 0
        end end
        local.get $acc call $log)"#, "6" },
    test_unreachable_after_return_not_hit => { r#"(func (export "_start")
        i32.const 42 call $log return unreachable)"#, "42" },
}
