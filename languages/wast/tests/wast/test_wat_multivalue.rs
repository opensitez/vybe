//! Multi-value proposal: blocks and functions that consume and produce more
//! than one stack value, and block parameters.
use crate::wat_exec;

wat_exec! {
    test_func_returns_two_values => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $pair (result i32 i32) i32.const 11 i32.const 22)
        (func (export "_start") call $pair call $log call $log))"#, "22" },
    test_block_result_value => { r#"(func (export "_start")
        block (result i32) i32.const 5 i32.const 6 i32.add end call $log)"#, "11" },
    test_block_param_consumed => { r#"(func (export "_start")
        i32.const 3 block (param i32) (result i32) i32.const 4 i32.add end call $log)"#, "7" },
    test_if_multi_value_result => { r#"(func (export "_start")
        i32.const 1 if (result i32) i32.const 100 else i32.const 200 end call $log)"#, "100" },
    test_loop_result_value => { r#"(func (export "_start")
        block (result i32) loop (result i32) i32.const 42 br 1 end end call $log)"#, "42" },
    test_block_two_results => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func (export "_start")
          block (result i32 i32) i32.const 7 i32.const 8 end i32.add call $log))"#, "15" },
    test_swap_via_multivalue_block => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func (export "_start")
          i32.const 10 i32.const 20
          block (param i32 i32) (result i32 i32) end
          i32.sub call $log))"#, "-10" },
    test_func_multi_return_used_in_arithmetic => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $divmod (param i32 i32) (result i32 i32)
          local.get 0 local.get 1 i32.div_u
          local.get 0 local.get 1 i32.rem_u)
        (func (export "_start")
          i32.const 17 i32.const 5 call $divmod i32.add call $log))"#, "5" },
}
