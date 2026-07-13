//! Conditional concepts — if/else, select, boolean logic built from integers,
//! guards, and value-selecting patterns (min/max/clamp/sign/abs).
use crate::wat_exec;

wat_exec! {
    test_if_true_branch => { r#"(func (export "_start")
        i32.const 1 if (result i32) i32.const 10 else i32.const 20 end call $log)"#, "10" },
    test_if_false_branch => { r#"(func (export "_start")
        i32.const 0 if (result i32) i32.const 10 else i32.const 20 end call $log)"#, "20" },
    test_select_picks_first => { r#"(func (export "_start")
        i32.const 111 i32.const 222 i32.const 1 select call $log)"#, "111" },
    test_select_picks_second => { r#"(func (export "_start")
        i32.const 111 i32.const 222 i32.const 0 select call $log)"#, "222" },
    test_boolean_and => { r#"(func (export "_start")
        i32.const 1 i32.const 1 i32.and call $log)"#, "1" },
    test_boolean_or => { r#"(func (export "_start")
        i32.const 0 i32.const 1 i32.or call $log)"#, "1" },
    test_boolean_not_via_eqz => { r#"(func (export "_start")
        i32.const 0 i32.eqz call $log)"#, "1" },
    test_short_circuit_and_pattern => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $both (param $a i32) (param $b i32) (result i32)
          local.get $a if (result i32) local.get $b else i32.const 0 end)
        (func (export "_start") i32.const 1 i32.const 5 call $both call $log))"#, "5" },
    test_min_of_two => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $min (param $a i32) (param $b i32) (result i32)
          local.get $a local.get $b local.get $a local.get $b i32.lt_s select)
        (func (export "_start") i32.const 8 i32.const 3 call $min call $log))"#, "3" },
    test_max_of_two => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $max (param $a i32) (param $b i32) (result i32)
          local.get $a local.get $b local.get $a local.get $b i32.gt_s select)
        (func (export "_start") i32.const 8 i32.const 3 call $max call $log))"#, "8" },
    test_abs_via_conditional => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $abs (param $x i32) (result i32)
          local.get $x i32.const 0 i32.lt_s
          if (result i32) i32.const 0 local.get $x i32.sub else local.get $x end)
        (func (export "_start") i32.const -42 call $abs call $log))"#, "42" },
    test_sign_of_number => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $sign (param $x i32) (result i32)
          local.get $x i32.const 0 i32.gt_s
          if (result i32) i32.const 1
          else local.get $x i32.const 0 i32.lt_s
               if (result i32) i32.const -1 else i32.const 0 end end)
        (func (export "_start") i32.const -7 call $sign call $log))"#, "-1" },
    test_clamp_to_range => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $clamp (param $x i32) (result i32)
          local.get $x i32.const 100 i32.gt_s
          if (result i32) i32.const 100
          else local.get $x i32.const 0 i32.lt_s
               if (result i32) i32.const 0 else local.get $x end end)
        (func (export "_start") i32.const 150 call $clamp call $log))"#, "100" },
    test_nested_if_grading => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $grade (param $s i32) (result i32)
          local.get $s i32.const 90 i32.ge_s
          if (result i32) i32.const 4
          else local.get $s i32.const 80 i32.ge_s
               if (result i32) i32.const 3
               else local.get $s i32.const 70 i32.ge_s
                    if (result i32) i32.const 2 else i32.const 1 end end end)
        (func (export "_start") i32.const 85 call $grade call $log))"#, "3" },
    test_select_with_computed_condition => { r#"(func (export "_start")
        i32.const 100 i32.const 200 i32.const 6 i32.const 4 i32.gt_s select call $log)"#, "100" },
    test_even_or_odd => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $parity (param $n i32) (result i32) local.get $n i32.const 1 i32.and)
        (func (export "_start") i32.const 13 call $parity call $log))"#, "1" },
}
