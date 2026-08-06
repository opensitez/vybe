//! Scope concepts — local vs global scope, per-call local isolation, label
//! scoping in nested blocks, and parameter shadowing behaviour.
use crate::wat_exec;

wat_exec! {
    test_locals_are_per_call => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $f (param $p i32) (result i32) (local $x i32)
          local.get $p i32.const 1 i32.add local.set $x local.get $x)
        (func (export "_start")
          i32.const 10 call $f drop i32.const 20 call $f call $log))"#, "21" },
    test_global_shared_across_calls => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func $bump global.get $g i32.const 1 i32.add global.set $g)
        (func (export "_start") call $bump call $bump call $bump global.get $g call $log))"#, "3" },
    test_local_does_not_leak_to_caller => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $inner (result i32) (local $x i32) i32.const 99 local.set $x local.get $x)
        (func (export "_start") (local $x i32)
          i32.const 5 local.set $x call $inner drop local.get $x call $log))"#, "5" },
    test_nested_block_labels_distinct => { r#"(func (export "_start")
        (local $r i32)
        block $outer block $inner
          i32.const 1 br_if $inner
          i32.const 0 local.set $r br $outer
        end i32.const 42 local.set $r end
        local.get $r call $log)"#, "42" },
    test_inner_label_shadows_by_depth => { r#"(func (export "_start")
        block block
          i32.const 7 call $log br 0 i32.const 8 call $log
        end i32.const 9 call $log end)"#, "7" },
    test_parameter_visible_throughout_function => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $f (param $p i32) (result i32)
          block local.get $p i32.const 100 i32.gt_s if unreachable end end
          local.get $p i32.const 2 i32.mul)
        (func (export "_start") i32.const 21 call $f call $log))"#, "42" },
    test_global_read_in_function => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (global $base i32 (i32.const 1000))
        (func $offset (param $d i32) (result i32) global.get $base local.get $d i32.add)
        (func (export "_start") i32.const 23 call $offset call $log))"#, "1023" },
    test_recursion_each_frame_own_locals => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $depth (param $n i32) (result i32) (local $marker i32)
          local.get $n local.set $marker
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $depth
               local.get $marker i32.add end)
        (func (export "_start") i32.const 4 call $depth call $log))"#, "10" },
    test_two_functions_independent_locals => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $a (result i32) (local $x i32) i32.const 1 local.set $x local.get $x)
        (func $b (result i32) (local $x i32) i32.const 2 local.set $x local.get $x)
        (func (export "_start") call $a call $b i32.add call $log))"#, "3" },
}
