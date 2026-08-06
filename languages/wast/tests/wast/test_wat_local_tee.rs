//! local.tee, local.get/set interplay, and multiple local slots.
use crate::wat_exec;

wat_exec! {
    test_local_tee_returns_and_stores => { r#"(func (export "_start")
        (local $x i32) i32.const 42 local.tee $x drop local.get $x call $log)"#, "42" },
    test_local_tee_value_used_immediately => { r#"(func (export "_start")
        (local $x i32) i32.const 5 local.tee $x i32.const 3 i32.add call $log)"#, "8" },
    test_local_tee_chained => { r#"(func (export "_start")
        (local $a i32) (local $b i32)
        i32.const 10 local.tee $a local.set $b
        local.get $a local.get $b i32.add call $log)"#, "20" },
    test_multiple_locals_independent => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func (export "_start")
          (local $a i32) (local $b i32) (local $c i32)
          i32.const 1 local.set $a i32.const 2 local.set $b i32.const 3 local.set $c
          local.get $a local.get $b i32.add local.get $c i32.add call $log))"#, "6" },
    test_local_default_is_zero => { r#"(func (export "_start")
        (local $x i32) local.get $x call $log)"#, "0" },
    test_local_f64_default_zero => { r#"(func (export "_start")
        (local $x f64) local.get $x call $log_f64)"#, "0.0" },
    test_local_reused_across_iterations => { r#"(func (export "_start")
        (local $i i32) (local $p i32)
        i32.const 1 local.set $p i32.const 5 local.set $i
        block loop
          local.get $i i32.eqz br_if 1
          local.get $p i32.const 2 i32.mul local.set $p
          local.get $i i32.const 1 i32.sub local.set $i br 0
        end end
        local.get $p call $log)"#, "32" },
    test_param_and_local_mix => { r#"(module
        (import "web:console" "log" (func $log (param i32)))
        (func $f (param $p i32) (result i32)
          (local $tmp i32) local.get $p i32.const 100 i32.add local.set $tmp local.get $tmp)
        (func (export "_start") i32.const 5 call $f call $log))"#, "105" },
    test_local_tee_i64 => { r#"(func (export "_start")
        (local $x i64) i64.const 9000000000 local.tee $x drop local.get $x call $log_i64)"#, "9000000000" },
}
